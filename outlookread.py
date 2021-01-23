from win32com.client import Dispatch
from time import sleep

outlook = Dispatch("Outlook.Application")

inbox = outlook.GetNamespace("MAPI").GetDefaultFolder(6) # "6" refers to the index of a folder - in this case,
# the inbox. You can change that number to reference
# any other folder
filter_by_string = "urn:schemas:mailheader:subject like '%CSS Workflow Request%'"
filter_by_date ="%last7days("+"urn:schemas:httpmail:datereceived"+")%"

search = outlook.AdvancedSearch( inbox, f"{filter_by_string} AND {filter_by_date}", True, "MySearch")

while search.Results.Count <= 0:
    pass
for item in search.Results:
    print(item.body)
