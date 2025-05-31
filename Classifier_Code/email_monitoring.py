import time
import os
import win32com.client
import io
from pdf2image.exceptions import PDFPageCountError
from pdf2image import convert_from_bytes
from Classifier_Code.Image_Training.loaded_model import invoice_predictor
from html_to_pdf import html_to_pdf, merge_pdfs


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

# Where  you want the result sent.
for_processing_path = r""
while True:
    # Don't know why, but ap_folder needs Item, and not Items either.
    ap_folder = namespace.Folders.Item("Accounts Payable")
    ap_team_folder_item = ap_folder.Folders.Item("AP TEAM")
    ap_inbox_folder_item = ap_folder.Folders.Item("Inbox")
    ap_completed_folder = ap_team_folder_item.Folders.Item("AP Completed")

    # faculty_approvals_folder = ap_inbox_folder_item.Folders.Item("Faculty Exp. Approvals - Faculty Mgmt. Profiles")


    # You need .Items in order for the messages to be recognized as such. Otherwise, I think its
    # treated as a folder object.
    # BUT you need to put the .Items in your iteration otherwise it won't be enumerable

    # last_ap_message = ap_inbox_folder_item.Items.GetLast().attachments
    last_ap_folder = ap_inbox_folder_item.Items
    items_in_folder = ap_inbox_folder_item.Items.count



    time.sleep(1)

    for message in last_ap_folder:
        try:
            if "Pythonic Classifier" in message.Categories:
                pass
            else:
                inbox_item_attachments = message.Attachments
                if message.Unread is True and inbox_item_attachments.Count > 0:

                    for index, message_attachments in enumerate(inbox_item_attachments):
                        position = index
                        if str(message_attachments).endswith(".pdf") is False:
                            pass
                        else:
                            print(index,message_attachments)

                            attachment = message_attachments
                            if str(attachment).endswith(".pdf"):
                                properties = attachment.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x37010102")
                                binary = io.BytesIO(properties)
                                value = None
                                # If you have pdfs that are password protected, remember to set the userpw parameter.
                                pdf_pass = ""
                                try:

                                    images = convert_from_bytes(binary.getvalue(), poppler_path=r"external_items\poppler-24.08.0\Library\bin",userpw=pdf_pass)
                                    value = invoice_predictor(images[0])
                                except PDFPageCountError:
                                    print("Password issue")

                                    pass

                                if value:
                                    attachment_file_path = os.path.join(for_processing_path, f"NOT READY ATTACHMENT {value}-{str(attachment)}")
                                    attachment.saveasfile(attachment_file_path)
                                    time.sleep(1)
                                    email_body = message.htmlbody
                                    email_body_path = os.path.join(for_processing_path, f"NOT READY email_body {value}-{str(attachment)}")
                                    try:
                                        html_to_pdf(email_body,email_body_path)
                                    except OSError as e:
                                        print(e)
                                        pass
                                    message.Unread = False
                                    message.Categories = "Pythonic Classifier"
                                    message.save()
                                    merged_path = os.path.join(for_processing_path, f"{value}-merged-{str(attachment)}")
                                    merge_pdfs(attachment_file_path,email_body_path,merged_path,pdf_pass=pdf_pass)

                                    os.remove(email_body_path)
                                    os.remove(attachment_file_path)

                    pass

                else:
                    pass

        except BaseException as e:
            # WHat likely happens here is that its trying to perform the checks while its being moved to another folder.
            print(f"Email {message.subject} caused an error {e}.")