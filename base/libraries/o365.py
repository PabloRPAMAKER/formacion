from O365 import Account, Connection,  MSGraphProtocol, Message


def credenciales_o365 (ruta:str, user:str, key:str, tenant:str, email:str):
    
    credentials = (user, key)
    # the default protocol will be Microsoft Graph
    scopes = ['https://graph.microsoft.com/Mail.ReadWrite', 'https://graph.microsoft.com/Mail.Send']


    account = Account(credentials, auth_flow_type='credentials', tenant_id=tenant)
    if account.authenticate():
        mailbox = account.mailbox(resource=email)
        inbox = mailbox.inbox_folder()
        inbox_bender = inbox.get_folder(folder_name='CERTIFICADOS')
        print (inbox_bender)
        for message in inbox_bender.get_messages(limit=30, download_attachments=True):
            if not message.is_read:
                print ('no leido')
                print (message.subject)
                print (message.has_attachments)
                message.attachments.download_attachments()
                print (message.attachments.__len__())
                adjunto = message.attachments.__getitem__(0)
                print (adjunto)
                for att in message.attachments:
                     print (att)
                     att.save(ruta)
                     return att
  
