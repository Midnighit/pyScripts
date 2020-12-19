from exiles_api import *
from datetime import datetime, timedelta

# save current time
now = datetime.utcnow()

conn = engines["gamedb"].connect()
conn_spl = engines["usersdb"].connect()

print("Run some SQL commands...")
# print("GAME_DB_URI:", GAME_DB_URI)

# MOVE CHAR TO NEW GUILD:
# Characters.move_to_guild(char_id=105, guild_id=1654940)

# GIVE OWNERSHIP OF THRALL WITH GIVEN OBJECT_ID TO NEW GUILD OR CHAR WITH GIVEN OWNER_ID
# Properties.give_thrall(object_id=1817012, owner_id=1655087)

# GRANT CLAIMS WITH OR WITHOUT COORDINATES WITH THE FORMAT ((x_min, x_max), (y_min, y_max), [(z_min, z_max)])
# Buildings.give_to_owner(old_owner_id=2058114, new_owner_id=1655087)
# Buildings.give_to_owner(old_owner_id=2207135, new_owner_id=1655087, loc=((-35000, -15000), (30000, 50000)))

# COPIES THE STATS OF A WEAPON OR ARMOR WITHIN A GIVEN INVENTORY (OWNER_ID) TO ALL EXISTING ONES OF THE SAME TYPE
# E.G. COPY AN UNMODIFIED YOG'S TOUCH OVER ALL MODDED OR UNNERFED YOG'S TOUCH
# ItemInventory.copy_stats(template_id=50523, owner_id=105)

# CHANGE A CHARACTER TIMESTAMP FOR THE RUINS SCRIPT. IF NO DATE IS GIVEN ASSUMES NOW.
# Characters.set_last_login(char_id=105)
# Characters.set_last_login(char_id=105, date=datetime.utcnow() + timedelta(days=30))

# REMOVE BUILDING PIECES BY COORDINATES AND OWNER:
# Buildings.remove_by_owner(owner_id=2256144)
# Buildings.remove_by_owner(owner_id=1655087, loc=((100, 200), (300, 400)))

# REMOVE ITEMS/EMOTES/FEATS FROM CONTAINER/INVENTORY IN THE GAME
# ItemInventory.remove(template_ids=(18400, 18401, 19605, 19606, 53001))
# ItemInventory.remove(template_ids=18400)

# REMOVE CHARACTERS FROM THE DB. ALSO REMOVES CLAN IF CHAR IS LAST MEMBER AND ACCOUNT IF PLAYER HAS NO OTHER CHARS
# Characters.remove(character_ids=(105, 106))
# Characters.remove(character_ids=105)

# REMOVE AND RESTORE AN OWNERS BUILDINGS FROM A BACKUP (REQUIRES backup.db IN SAVED FOLDER)
Buildings.restore_from_backup(owner_id=2256144)
# Buildings.restore_from_backup(1655087, loc=((-150000, -20000), (100000, 180000)))

# WORKING WITH THE SQLALCHEMY ORM AND EXILES_API
# char = session.query(Characters).get(105)
# guild = char.guild
# guild.name = "Some other name"
# funcom_id = char.user.funcom_id
# members = guild.members

# MISC SQL STATEMENTS
# conn.execute("UPDATE guilds SET name = 'Savanna Arena' WHERE guildId = 14")
# conn_spl.execute("DELETE FROM users WHERE funcom_id = 8187A5834CD94E58")

execTime = datetime.utcnow() - now
execTimeStr = str(execTime.seconds) + "." + str(execTime.microseconds)
print(f"Done!\nRequired time: {execTimeStr} sec.")
