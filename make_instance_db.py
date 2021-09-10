from exiles_api import *

# save current time
now = datetime.utcnow()

print("Copying to destination db...")

"""
Description:
make_instance_db is meant as an easy way to take all the usual steps for creating a new db instance for raids, events etc.
As input it requires a live db to copy from (usually game.db), an empty destination db and additional filter arguments
to setup which parts of the live db should be copied over.
All arguments have default values and can be left empty if the default should be used.

Arguments:
source_db:
    Filename of the source db file to copy from. Directory is assumed to be the CE saved folder where the game.db resides
    default: "game.db"
dest_db:
    Filename of the destination db file to copy from. Directory is assumed to be the CE saved folder where the game.db resides
    default: "dest.db"
owner_ids:
    Either chararacters.id or guilds.guildId of the characters, guilds and their buildings that should be copied.
    If owner_id is None, will copy all characters and guilds respectively.
    Argument can be given as a number or a comma separated list encapsulated in [] or ().
    dafault: None
loc:
    Given in the format [(min_x, max_x), (min_y, max_y), (min_z, max_z)] for coordinates x, y and z.
    The z coordinate is optional and can be omitted. The min value has to be smaller or equal to the max value.
    If loc is None, will copy buildings everywhere.
    default: None
with_chars:
    If set to True will automatically copy all characters for guilds contained in owner_ids as well.
    default: True
with_alts:
    Copies all characters belonging to accounts for any one of the targeted characters as well.
    default: True
inverse_owners:
    If inverse_owners is set to True, will instead copy owners and buildings owned by characters/guilds that are not given and
    outside of coordinates given.
    default: False
mod_names:
    The names of the mods whose config and mod controllers are copied. The names are those of the .pak files in the modlist.
    Note: mod related settings that are attached to objects in the world or characters (i.e. character stats for RoleplayRedux)
    will be copied with those regardless of the mod_names argument.
    If mod_names is None, will copy all mods present in the source db.
    default: None
inverse_mods:
    If inverse_mods is set to True, will instead copy mods that are not given.
    default=False
"""

make_instance_db(
    owner_ids=[2496381, 2945124, 3012565, 2078446, 2982733, 22, 1654923, 1654980, 2993719, 2901408, 2396885, 2849597, 2213918, 3009749, 107],
    mod_names=["Roleplay", "RavencrestCouriers"], 
    inverse_mods=True
)

execTime = datetime.utcnow() - now
execTimeStr = str(execTime.seconds) + "." + str(execTime.microseconds)
print(f"Done!\nRequired time: {execTimeStr} sec.")
