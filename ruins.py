import sys
import random
import logging
from time import time
from datetime import datetime, timedelta
from logger import get_logger
from config import (
    LOG_LEVEL_STDOUT, LOG_LEVEL_FILE, RUINS_CLAN_ID, INACTIVITY, LONG_INACTIVE, EVENT_LOG_HOLD_BACK,
    OBJECT_LIMITS, OWNER_WHITELIST, ALLOWANCE_INCLUDES_INACTIVES, PURGE
)
from exiles_api import (
    StaticBuildables, session, engines, Guilds, GameEvents, ActorPosition, Buildings, Tiles, Characters,
    DeleteChars, OwnersCache, ObjectsCache, Properties, Thralls, BuildableHealth
)

# catch unhandled exceptions
logger = get_logger('ruins.log', log_level_stdout=LOG_LEVEL_STDOUT, log_level_file=LOG_LEVEL_FILE)


def handle_exception(exc_type, exc_value, exc_traceback):
    if issubclass(exc_type, KeyboardInterrupt):
        sys.__excepthook__(exc_type, exc_value, exc_traceback)
        return

    logger.error("Uncaught exception", exc_info=(exc_type, exc_value, exc_traceback))


sys.excepthook = handle_exception

# save current time
start_time = time()
now = datetime.utcnow()
if LOG_LEVEL_STDOUT > logging.INFO:
    print("Executing ruins script...")
logger.info("Executing ruins script...")

# make sure a ruins clan exists
ruins_clan = session.query(Guilds).get(RUINS_CLAN_ID)
if not ruins_clan:
    ruins_clan = Guilds(id=RUINS_CLAN_ID, name='Ruins')
    session.add(ruins_clan)

# seed the random number generator
random.seed()

# store some timestamps required for following operations
now_ts = int(now.timestamp())
ia_ts = int((now - INACTIVITY).timestamp())
lia_ts = int((now - LONG_INACTIVE).timestamp())
el_ts = int((now - EVENT_LOG_HOLD_BACK).timestamp())

""" Cull the event log """
logger.debug("Culling the events log.")
session.query(GameEvents).filter(GameEvents.world_time <= el_ts).delete()

""" Delete some blacklisted items placed in the world """
logger.debug("Deleting blacklisted items.")
for limit in OBJECT_LIMITS:
    # non-existing per_member value is assumed to be 0
    per_member = limit['pm'] if 'pm' in limit else 0
    # construct the filter consisting of the link between the two tables and the 'like' filters over the classes
    filter_link = (ActorPosition.id == Buildings.object_id) & Buildings.owner_id.notin_(OWNER_WHITELIST)
    filter_class = None
    for class_name in limit['obj']:
        if filter_class is None:
            filter_class = ActorPosition.class_.like(f"%{class_name}%")
        else:
            filter_class = filter_class | ActorPosition.class_.like(f"%{class_name}%")
    # consolidate all objects belonging to a given owner i.e. {owner_id: [object1, object2, object3...]}
    owners = {}
    for id, building in session.query(ActorPosition, Buildings).filter(filter_link & (filter_class)).all():
        if building.owner is None:
            continue
        elif building.owner not in owners:
            owners[building.owner] = [id]
        else:
            owners[building.owner] += [id]
    # for each owner check if number of objects found exceeds the allowance
    for owner, objects in owners.items():
        if ALLOWANCE_INCLUDES_INACTIVES:
            num_members = len(owner.members) if owner.is_guild else 1
        else:
            num_members = len(owner.active_members(INACTIVITY)) if owner.is_guild else 1
        diff = len(objects) - limit['max'] - num_members * per_member
        # if allowance has been exceeded
        if diff > 0:
            logger.info(f"Deleting the following objects from {owner.name} ({owner.id}):")
            # pick diff amount of object_ids from the objects list and remove them from the db
            picked_objects = random.sample(objects, diff)
            picked_object_ids = []
            for obj in picked_objects:
                picked_object_ids.append(obj.id)
                tp = f"TeleportPlayer {round(obj.x)} {round(obj.y)} {round(obj.z)}"
                logger.info(f"{obj.class_[obj.class_.rfind('.')+1:]} ({obj.id}): {tp}")
            Tiles.remove(picked_object_ids, autocommit=False)

""" Delete old characters from the db and clean up behind them """
logger.debug("Deleting old characters.")
char_ids = set()
player_ids = {}
# get all characters who logged in before a configured time
filter = (Characters._last_login <= lia_ts) & Characters.id.notin_(OWNER_WHITELIST)
deleted_chars = tuple(session.query(Characters).filter(filter).all())
for char in deleted_chars:
    user = char.user
    if user:
        player = f"{user.disc_user} ({user.disc_id}) with FuncomID {user.funcom_id} and PlayerID {char.player_id}"
    else:
        player = f" with FuncomID {char.account.funcom_id} and PlayerID {char.player_id}"
    logger.info(f"Deleting {char.name} ({char.id}) belonging to player {player}.")
    player_ids.update({char.player_id: char.name})
    char_ids.add(char.id)
# use Characters.remove as opposed to session.delete(char) to not just remove it from the characters table
Characters.remove(char_ids, autocommit=False, whitelist=OWNER_WHITELIST)
# Log all chars that have been delete this way so they can be removed in TERPO too when the server runs
DeleteChars.add(player_ids, autocommit=False)

""" Rename applicable characters/guilds to ruins or rename them back to their original names"""
# update the OwnersCache table and load it into a dict for convenient lookup
logger.debug("Renaming characters and guilds from and to 'Ruins'.")
OwnersCache.update(RUINS_CLAN_ID, autocommit=False)
ownerscache = {id: n for id, n in session.query(OwnersCache.id, OwnersCache.name).all()}

# go through all Chars/Guilds named 'Ruins' that are (no longer) inactive and rename them to their original name
for char in session.query(Characters).filter((Characters.name == 'Ruins') & (Characters._last_login > ia_ts)).all():
    if char.id in ownerscache:
        logger.info(f"Renaming char with id {char.id} back from 'Ruins' to '{ownerscache[char.id]}'.")
        char.name = ownerscache[char.id]
for guild in session.query(Guilds).filter((Guilds.name == 'Ruins') & (Guilds.id != RUINS_CLAN_ID)).all():
    if not guild.is_inactive(INACTIVITY) and guild.id in ownerscache:
        logger.info(f"Renaming guild with id {guild.id} back from 'Ruins' to '{ownerscache[guild.id]}'.")
        guild.name = ownerscache[guild.id]

# go through all characters that are not whitelisted, not in a guild and inactive
filter = (Characters.id.notin_(OWNER_WHITELIST)) & (Characters.guild_id.is_(None)) & (Characters._last_login <= ia_ts)
for char in session.query(Characters).filter(filter).order_by(Characters.id).all():
    # do function call once and store the result to save time
    has_tiles = char.has_tiles()
    # if char is named Ruins but has no buildings left, rename back to original name in case they return
    if char.name == 'Ruins' and not has_tiles and char.id in ownerscache:
        logger.info(f"Renaming char with id {char.id} back from 'Ruins' to '{ownerscache[char.id]}'.")
        char.name = ownerscache[char.id]
    # if char is not named Ruins and still has buildings left, rename to ruins
    elif char.name != 'Ruins' and has_tiles:
        logger.info(f"Renaming char with id {char.id} from '{ownerscache[char.id]}' to 'Ruins'.")
        char.name = 'Ruins'

# go through all guilds since filtering for inactivity is more complicated there
filter = (Guilds.id.notin_(OWNER_WHITELIST)) & (Guilds.id != RUINS_CLAN_ID)
for guild in session.query(Guilds).filter(filter).order_by(Guilds.id).all():
    if guild.is_inactive(INACTIVITY):
        # do function call once and store the result to save time
        has_tiles = guild.has_tiles()
        # if guild has no members, delete it. This will place it's buidlings on the ObjectsCache when next updated
        if len(guild.members) == 0:
            session.delete(guild)
        # if guild is named Ruins but has no buildings left, rename back to original name in case the owner(s) return
        elif guild.name == 'Ruins' and not has_tiles and guild.id in ownerscache:
            logger.info(f"Renaming guild with id {guild.id} back from 'Ruins' to '{ownerscache[guild.id]}'.")
            guild.name = ownerscache[guild.id]
        # if guild is not named Ruins and still has buildings left, rename to ruins
        elif guild.name != 'Ruins' and has_tiles:
            logger.info(f"Renaming guild with id {guild.id} from '{ownerscache[guild.id]}' to 'Ruins'.")
            guild.name = 'Ruins'

""" move all owner_id 0 objects to the dedicated ruins clan """
reassigned = []
stat_builds = session.query(StaticBuildables.id).scalar_subquery()
for tile in session.query(Buildings).filter(Buildings.object_id.notin_(stat_builds) & (Buildings.owner_id == 0)):
    reassigned.append(tile.object_id)
    tile.owner_id = RUINS_CLAN_ID

if len(reassigned) > 0:
    logger.info(f"Moving owner_id 0 objects to dedicated Ruins clan: {str(reassigned)}.")

""" move all ownerless objects to the dedicated ruins clan """
logger.debug("Moving ownerless objects to dedicated ruins clan.")
# update the ObjectsCache table and load it into a dict for convenient lookup
ObjectsCache.update(RUINS_CLAN_ID, autocommit=False)
objectscache = {id: ts for id, ts in session.query(ObjectsCache.id, ObjectsCache._timestamp).all()}

# add all object_ids that belong to the dedicated ruins clan into a set for an easy check
ruins_clan_query = session.query(Buildings.object_id).filter_by(owner_id=RUINS_CLAN_ID)
ruins_clan = set(r for r, in ruins_clan_query.all())

# go through all the objects that either have no owner or are in the dedicated ruins clan
reassigned = []
for object_id, object_ts in objectscache.items():
    # if ownerless object is not yet in the ruins clan, move it there
    if object_id not in ruins_clan:
        # since ObjectsCache was just updated, it should only contain existing objects
        reassigned.append(object_id)
        obj = session.query(Buildings).filter_by(object_id=object_id).one()
        obj.owner_id = RUINS_CLAN_ID

if len(reassigned) > 0:
    logger.info(f"Moving no owner objects to dedicated Ruins clan: {str(reassigned)}.")

""" damage or remove buildins belonging to 'Ruins' owners """
logger.debug("Applying damage to and removing buildings belonging to the dedicated ruins clan.")
# index all thralls by their owners so they can be removed alongside their owners buildings
thralls = {}
for property in session.query(Properties).filter(Properties.name.like("%OwnerUniqueID")).all():
    owner_id = property.owner_id
    if owner_id in thralls:
        thralls[owner_id] += [property.object_id]
    else:
        thralls[owner_id] = [property.object_id]

chars_query = session.query(Characters.id).filter_by(name='Ruins')
guilds_query = session.query(Guilds.id).filter_by(name='Ruins')
filter = Buildings.owner_id.in_(chars_query.union(guilds_query)) & Buildings.owner_id.notin_(OWNER_WHITELIST)
ruins_query = session.query(Buildings).filter(filter)
damaged, removed, killed = [], [], []
for building in ruins_query.all():
    # if building has no owner, the time since last login needs to be determined via objectscache timestamp
    if building.object_id in objectscache:
        time_since_inactive = timedelta(seconds=now_ts-objectscache[building.object_id])
    # If owner had just been deleted for inactivity assume that their buildings should also be removed
    elif building.owner in deleted_chars:
        time_since_inactive = INACTIVITY + PURGE
    # otherwise use the timestamp of the last login
    elif building.owner.last_login is not None:
        time_since_inactive = now - building.owner.last_login - INACTIVITY
    # should never get here but catch it anyway
    else:
        logger.warning(f"building.owner.last_login was None for {building} when damage was supposed to be calculated.")
    # calculate the damage percentage based on the time since the owner became inactive relative to the PURGE timedelta
    dmg = 1 - time_since_inactive / PURGE
    # if damage >= 100% simply remove the objects from db
    if dmg <= 0:
        removed.append(building.object_id)
        Tiles.remove(building.object_id, autocommit=False)
        # if owner had thralls remove those too
        if building.owner_id in thralls:
            killed += thralls[building.owner_id]
            Thralls.remove(thralls[building.owner_id], autocommit=False)
            del thralls[building.owner_id]
    # if damage < 100% damage and that part of the object isn't already more damaged than that, damage it
    else:
        damaged.append(building.object_id)
        for part in session.query(BuildableHealth).filter_by(object_id=building.object_id).all():
            if part.health_percentage > dmg:
                part.health_percentage = dmg

if len(damaged) > 0:
    logger.info(f"Damaging objects: {str(damaged)}.")
if len(removed) > 0:
    logger.info(f"Removing objects: {str(removed)}.")
if len(killed) > 0:
    logger.info(f"Killing thralls: {str(killed)}.")

session.commit()
logger.debug("Cleaning up the db.")
for name, engine in engines.items():
    with engine.connect() as conn:
        conn.execute('VACUUM')
        conn.execute('REINDEX')
        conn.execute('ANALYZE')
        conn.execute('PRAGMA integrity_check')
        # If there are any pending changes, commit them
        if conn.in_transaction():
            conn.execute('COMMIT')

exec_time = time() - start_time
if LOG_LEVEL_STDOUT > logging.INFO:
    print(f"Done! Required time: {exec_time:.3f} sec.")
logger.info(f"Done! Required time: {exec_time:.3f} sec.")
