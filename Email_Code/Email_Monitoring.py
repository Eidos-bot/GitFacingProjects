import win32com.client
import io
from pdf2image import convert_from_bytes

outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")
inbox = namespace.GetDefaultFolder(6)
sent = namespace.GetDefaultFolder(5)
outbox = namespace.GetDefaultFolder(4)
deleted = namespace.GetDefaultFolder(3)

# "6" refers to the index of a folder - in this case,
# the inbox. You can change that number to reference
# any other folder

# messages = inbox.Items
# all_messages = list(messages)
# print(len(all_messages))
# body_content = message.body
# print(inbox)

# Don't know why, but ap_folder needs Item, and not Items either.
ap_folder = namespace.Folders.Item("Accounts Payable")
ap_team_folder_item = ap_folder.Folders.Item("AP TEAM")
ap_inbox_folder_item = ap_folder.Folders.Item("Inbox")
ap_completed_folder = ap_team_folder_item.Folders.Item("AP Completed")

# faculty_approvals_folder = ap_inbox_folder_item.Folders.Item("Faculty Exp. Approvals - Faculty Mgmt. Profiles")


# You need .Items in order for the messages to be recognized as such. Otherwise, I think its
# treated as a folder object.
# BUT you need to put the .Items in your iteration otherwise it won't be enumerable

last_faculty_message = ap_inbox_folder_item.Items.GetLast().attachments

for i in range(last_faculty_message.count):
    n = i+1
    attachment = last_faculty_message.item(n)
    if str(attachment).endswith(".pdf"):
        properties = attachment.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x37010102")
        binary = io.BytesIO(properties)
        image = convert_from_bytes(binary.getvalue())
        print(binary)