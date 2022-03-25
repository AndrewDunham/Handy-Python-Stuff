from appscript import app, k

outlook = app('Microsoft Outlook')

def sendEmail(residentName, residentEmail):
    msg = outlook.make(
        new=k.outgoing_message,
        with_properties={
            k.subject: 'Email Subject',
            k.plain_text_content: 'Email body'})

    msg.make(
        new=k.recipient,
        with_properties={
            k.email_address: {
                k.name: residentName,
                k.address: residentEmail}})

    msg.open()
    msg.activate()
    #msg.save()
