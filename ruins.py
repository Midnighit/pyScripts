import random
from config import *
from datetime import datetime, timedelta
from exiles_api import *

# save current time
now = datetime.utcnow()
print("Executing ruins script...")

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
session.query(GameEvents).filter(GameEvents.world_time <= el_ts).delete()

""" Delete some blacklisted items placed in the world """
for limit in OBJECT_LIMITS:
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
    for id, owner_id in session.query(ActorPosition.id, Buildings.owner_id).filter(filter_link & (filter_class)).all():
        if owner_id not in owners:
            owners[owner_id] = [id]
        else:
            owners[owner_id] += [id]
    # for each owner check if number of objects found exceeds the allowance
    for owner_id, objects in owners.items():
        diff = len(objects) - limit['max']
        # if allowance has been exceeded
        if diff > 0:
            # pick diff amount of object_ids from the objects list and remove them from the db
            Tiles.remove(random.sample(objects, diff), autocommit=False)

""" Delete old characters from the db and clean up behind them """
char_ids = set()
player_ids = {}
# get all characters who logged in before a configured time
filter = (Characters._last_login <= lia_ts) & Characters.id.notin_(OWNER_WHITELIST)
for char in session.query(Characters).filter(filter).all():
    player_ids.update({char.player_id: char.name})
    char_ids.add(char.id)
# use Characters.remove as opposed to session.delete(char) to not just remove it from the characters table
Characters.remove(char_ids, autocommit=False)
# Log all chars that have been delete this way so they can be removed in TERPO too when the server runs
DeleteChars.add(player_ids, autocommit=False)

""" Rename applicable characters/guilds to ruins or rename them back to their original names"""
# update the OwnersCache table and load it into a dict for convenient lookup
OwnersCache.update(RUINS_CLAN_ID, autocommit=False)
ownerscache = {id: n for id, n in session.query(OwnersCache.id, OwnersCache.name).all()}

# go through all Chars/Guilds named 'Ruins' that are (no longer) inactive and rename them to their original name
for char in session.query(Characters.id).filter((Characters.name=='Ruins') & (Characters._last_login > ia_ts)).all():
    if char.id in ownerscache:
        char.name = ownerscache[char.id]
for guild in session.query(Guilds).filter((Guilds.name=='Ruins') & (Guilds.id != RUINS_CLAN_ID)).all():
    if not guild.is_inactive(INACTIVITY) and guild.id in ownerscache:
        guild.name = ownerscache[guild.id]

# go through all characters that are not whitelisted, not in a guild and inactive
filter = (Characters.id.notin_(OWNER_WHITELIST)) & (Characters.guild_id == None) & (Characters._last_login <= ia_ts)
for char in session.query(Characters).filter(filter).order_by(Characters.id).all():
    # do function call once and store the result to save time
    has_tiles = char.has_tiles()
    # if char is named Ruins but has no buildings left, rename back to original name in case they return
    if char.name == 'Ruins' and not has_tiles and char.id in ownerscache:
        char.name = ownerscache[char.id]
    # if char is not named Ruins and still has buildings left, rename to ruins
    elif char.name != 'Ruins' and has_tiles:
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
            guild.name = ownerscache[guild.id]
        # if guild is not named Ruins and still has buildings left, rename to ruins
        elif guild.name != 'Ruins' and has_tiles:
            guild.name = 'Ruins'

""" move all ownerless objects to the dedicated ruins clan """
# update the ObjectsCache table and load it into a dict for convenient lookup
ObjectsCache.update(RUINS_CLAN_ID, autocommit=False)
objectscache = {id: ts for id, ts in session.query(ObjectsCache.id, ObjectsCache._timestamp).all()}

# add all object_ids that belong to the dedicated ruins clan into a set for an easy check
ruins_clan_query = session.query(Buildings.object_id).filter_by(owner_id=RUINS_CLAN_ID)
ruins_clan = set(r for r, in ruins_clan_query.all())

# go through all the objects that either have no owner or are in the dedicated ruins clan
for object_id, object_ts in objectscache.items():
    # if ownerless object is not yet in the ruins clan, move it there
    if object_id not in ruins_clan:
        # since ObjectsCache was just updated, it should only contain existing objects
        obj = session.query(Buildings).filter_by(object_id=object_id).one()
        obj.owner_id = RUINS_CLAN_ID

""" damage or remove buildins belonging to 'Ruins' owners """
# go through all buildings in BuildableHealth that belong to ruins
chars_query = session.query(Characters.id).filter_by(name='Ruins')
guilds_query = session.query(Guilds.id).filter_by(name='Ruins')
filter = Buildings.owner_id.in_(chars_query.union(guilds_query)) & Buildings.owner_id.notin_(OWNER_WHITELIST)
ruins_query = session.query(Buildings).filter(filter)
for building in ruins_query.all():
    # if building has no owner, the time since last login needs to be determined via objectscache timestamp
    if building.object_id in objectscache:
        time_since_inactive = timedelta(seconds=now_ts-objectscache[building.object_id])
    # otherwise use the timestamp of the last login
    else:
        time_since_inactive = now - building.owner.last_login - INACTIVITY
    # calculate the damage percentage based on the time since the owner became inactive relative to the PURGE timedelta
    dmg = 1 - time_since_inactive / PURGE
    # if damage >= 100% simply remove the objects from db
    if dmg <= 0:
        Tiles.remove(building.object_id, autocommit=False)
    # if damage < 100% damage and that part of the object isn't already more damaged than that, damage it
    else:
        for part in session.query(BuildableHealth).filter_by(object_id=building.object_id).all():
            if part.health_percentage > dmg:
                part.health_percentage = dmg

session.commit()

for name, engine in engines.items():
    with engine.connect() as conn:
        conn.execute('VACUUM')
        conn.execute('ANALYZE')

execTime = datetime.utcnow() - now
execTimeStr = str(execTime.seconds) + "." + str(execTime.microseconds)
print(f"Done!\nRequired time: {execTimeStr} sec.")
