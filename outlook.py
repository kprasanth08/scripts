from win32com.client import Dispatch
from _datetime import datetime

_t = datetime.now()

outlook = Dispatch("Outlook.Application").GetNamespace("MAPI")

inbox = outlook.GetDefaultFolder(6).Folders
# "6" refers to the index of a folder - in this case,
# the inbox. You can change that number to reference
# any other folder
count = 0
for folder in inbox:
    emails = folder.Items
    for email in emails:
        if 'CSS' in email.subject:
            # print(email.body)
            count += 1
            pass

inbox = outlook.GetDefaultFolder(6)
for email in inbox.Items:
    if 'CSS' in email.subject:
        # print(email.recipients)
        count += 1
        pass

print(f'end -{datetime.now() - _t} {count}')

_t = datetime.now()
outlook = Dispatch("Outlook.Application")

inbox = outlook.GetNamespace("MAPI").GetDefaultFolder(6)
filter_by_string = "urn:schemas:mailheader:subject like '%CSS%'"
filter_by_date = "%last7days(" + "urn:schemas:httpmail:datereceived" + ")%"

search = outlook.AdvancedSearch(inbox, f"{filter_by_string}", True, "MySearch")

while search.Results.Count <= 0:
    pass
for item in search.Results:
    # print(item.body)
    pass

print(f'end -{datetime.now() - _t} {search.Results.Count}')
