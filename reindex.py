import sys
from datetime import datetime
from config import LOG_LEVEL, RI_DEST, RI_SOURCE
from logger import get_logger
from exiles_api import engines, session, Properties

# catch unhandled exceptions
logger = get_logger('reindex.log', LOG_LEVEL)


def handle_exception(exc_type, exc_value, exc_traceback):
    if issubclass(exc_type, KeyboardInterrupt):
        sys.__excepthook__(exc_type, exc_value, exc_traceback)
        return

    logger.error("Uncaught exception", exc_info=(exc_type, exc_value, exc_traceback))


sys.excepthook = handle_exception

# save current time
now = datetime.utcnow()
logger.info("Starting the reindexing process:")

dest = RI_DEST
source = RI_SOURCE
existing_ids = (
    "SELECT DISTINCT id FROM actor_position UNION "
    "SELECT DISTINCT object_id AS id FROM buildings UNION "
    "SELECT DISTINCT char_id AS id FROM character_stats UNION "
    "SELECT DISTINCT id AS id FROM characters UNION "
    "SELECT DISTINCT guildId AS id FROM guilds"
)
owner_ids = "SELECT DISTINCT id AS id FROM characters UNION	SELECT DISTINCT guildId AS id FROM guilds"
source_idx = 0

with engines["gamedb"].begin() as conn:
    for dest_id in dest:
        # make sure dest_id isn't taken already
        if conn.execute(f"SELECT COUNT(id) FROM ({existing_ids}) WHERE id={dest_id}").scalar():
            logger.warning(f"Dest ID {dest_id} is already taken! Skipping this ID.")
            continue

        # if source at index idx is either None or doesn't exist at all, simply create an empty guild with dest_id
        if len(source) <= dest.index(dest_id) or source[dest.index(dest_id)] is None:
            logger.debug(f"creating empty guild 'Reserved' at ID {dest_id}.")
            conn.execute(f"INSERT INTO guilds VALUES ({dest_id}, 'Reserved', '', '', -1, -1, 0)")

        # if a source_id can be determined all occurances of source_id have to be changed to dest_id
        else:
            # make sure that source_id belongs to a character or clan (i.e. it belongs to an owner)
            source_id = source[dest.index(dest_id)]
            if not conn.execute(f"SELECT COUNT(id) FROM ({owner_ids}) WHERE id={source_id}").scalar():
                logger.debug(f"Source ID {source_id} is exists but is not a character or clan. Skipping this ID.")
                continue

            # execute all the simple relabeling
            logger.debug(f"Reindexing source ID {source_id} to dest ID {dest_id}.")
            conn.execute(f"UPDATE actor_position SET id = {dest_id} WHERE id = {source_id}")
            conn.execute(f"UPDATE buildings SET owner_id = {dest_id} WHERE owner_id = {source_id}")
            conn.execute(f"UPDATE character_stats SET char_id = {dest_id} WHERE char_id = {source_id}")
            conn.execute(f"UPDATE characters SET id = {dest_id} WHERE id = {source_id}")
            conn.execute(f"UPDATE characters SET guild = {dest_id} WHERE guild = {source_id}")
            conn.execute(f"UPDATE destruction_history SET owner_id = {dest_id} WHERE owner_id = {source_id}")
            conn.execute(f"UPDATE events SET member = {dest_id} WHERE member = {source_id}")
            conn.execute(f"UPDATE events SET guild = {dest_id} WHERE guild = {source_id}")
            conn.execute(f"UPDATE follower_markers SET owner_id = {dest_id} WHERE owner_id = {source_id}")
            conn.execute(f"UPDATE game_events SET ownerId = {dest_id} WHERE ownerId = {source_id}")
            conn.execute(f"UPDATE game_events SET ownerGuildId = {dest_id} WHERE ownerGuildId = {source_id}")
            conn.execute(f"UPDATE guilds SET guildId = {dest_id} WHERE guildId = {source_id}")
            conn.execute(f"UPDATE guilds SET owner = {dest_id} WHERE owner = {source_id}")
            conn.execute(f"UPDATE item_inventory SET owner_id = {dest_id} WHERE owner_id = {source_id}")
            conn.execute(f"UPDATE item_properties SET owner_id = {dest_id} WHERE owner_id = {source_id}")
            conn.execute(f"UPDATE properties SET object_id = {dest_id} WHERE object_id = {source_id}")
            conn.execute(f"UPDATE purgescores SET purgeid = {dest_id} WHERE purgeid = {source_id}")

            # give thralls and pets (if any) belonging to source_id to the dest_id instead
            object_ids = Properties.get_thrall_object_ids(owner_id=source_id)
            Properties.give_thrall(object_ids, dest_id, autocommit=False)

# try to commit all changes
session.commit()

execTime = datetime.utcnow() - now
execTimeStr = str(execTime.seconds) + "." + str(execTime.microseconds)
logger.info(f"Done! Required time: {execTimeStr} sec.")
