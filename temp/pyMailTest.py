from pyMail import pyMail
      
def main():
    # Create pyMail object
    mail = pyMail(input("Email: "), input("Password: "))
    mail.compose(input("Email Subjet: "), input("Email Body: "), input("Receiver Email: "))
    if(input("Attach file? (y/n): ") == 'y'):
        attachments = []
        attachments.append(input("File name: "))
        while attachAdditionalFile := input("Attach additional file? (y/n): ") == 'y':
            attachments.append(input("File name: "))
        mail.attach(attachments)
    mail.send()
    
if __name__ == "__main__":
    main()