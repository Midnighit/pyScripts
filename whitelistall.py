from config import *
from exiles_api import session, Users
from datetime import datetime

# save current time
now = datetime.utcnow()
print("Whitelisting everyone in the supplemental database...")

try:
    with open(WHITELIST_PATH, 'r') as f:
        lines = f.readlines()
# if file doesn't exist create an empty list
except:
    with open(WHITELIST_PATH, 'w') as f:
        pass
    lines = []

# split lines into id and name. Remove duplicates.
filtered = set()
names = {}
for line in lines:
    if line != "\n" and not "INVALID" in line:
        res = line.split(':')
        id = res[0].strip()
        if len(res) > 1:
            name = res[1].strip()
        else:
            name = 'Unknown'
        filtered.add(id)
        # if duplicate values exist, prioritize those containing a funcom_name
        if not id in names or names[id] == 'Unknown':
            names[id] = name

# go through the Users table and supplement missing users if any
for user in session.query(Users).all():
    if not user.funcom_id in filtered:
        filtered.add(user.funcom_id)
        names[user.funcom_id] = 'Unknown'

# create lines to write into new whitelist.txt
wlist = []
for id in filtered:
    wlist.append(id + ':' + names[id] + '\n')
    wlist.sort()

# overwrite / write the new file with the contenst of wlist
with open(WHITELIST_PATH, 'w') as f:
    f.writelines(wlist)

execTime = datetime.utcnow() - now
execTimeStr = str(execTime.seconds) + "." + str(execTime.microseconds)
print(f"Done!\nRequired time: {execTimeStr} sec.")
