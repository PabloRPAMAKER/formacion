from O365 import Account, Connection,  MSGraphProtocol, Message


def credenciales_o365 (ruta:str, user:str, key:str, tenant:str, email:str):
    #credentials = ('cee6b0cc-79ae-40bb-b58d-271882f0ce6c', 'LE18Q~ON6fFqed2u~bFVv3ArbWYq5hLqtK01WcQj')
    credentials = (user, key)
    # the default protocol will be Microsoft Graph
    scopes = ['https://graph.microsoft.com/Mail.ReadWrite', 'https://graph.microsoft.com/Mail.Send']

    #account = Account(credentials, auth_flow_type='credentials', tenant_id='11246493-edf9-411c-aa14-89d713937614')
    account = Account(credentials, auth_flow_type='credentials', tenant_id=tenant)
    if account.authenticate():
        # account.mailbox
        # m = account.new_message(resource='rpa.prl@limcamar.es')
        # m.to.add('pablo.saura@limcamar.es')
        # m.subject = 'Testing!'
        # m.body = "George Best quote: I've stopped drinking, but only while I'm asleep."
        # m.send()
        mailbox = account.mailbox(resource=email)
        #inbox = mailbox.get_folder(folder_name='Inbox')
        #print (inbox)

        inbox = mailbox.inbox_folder()
        inbox_bender = inbox.get_folder(folder_name='CERTIFICADOS')
        print (inbox_bender)
        for message in inbox_bender.get_messages(limit=10, download_attachments=True):
            if not message.is_read:
                print ('no leido')
                #message.mark_as_read()
                print (message.subject)
                print (message.has_attachments)
                message.attachments.download_attachments()
                print (message.attachments.__len__())
                adjunto = message.attachments.__getitem__(0)
                #adjunto.save('C:\Users\RPA\RPAMAKER\Dokify\base\output')
                print (adjunto)
                for att in message.attachments:
                     print (att)
                     att.save(ruta)
                     return att
  

def enviar_email (client_id:str, client_secret:str):
    credentials = ('cee6b0cc-79ae-40bb-b58d-271882f0ce6c', 'LE18Q~ON6fFqed2u~bFVv3ArbWYq5hLqtK01WcQj')
    account = Account(credentials, auth_flow_type='credentials', tenant_id='11246493-edf9-411c-aa14-89d713937614')
    account.mailbox
    m = account.new_message()
    m.to.add('pablo.saura@limcamar.es')
    m.subject = 'Testing!'
    m.body = "George Best quote: I've stopped drinking, but only while I'm asleep."
    m.send()