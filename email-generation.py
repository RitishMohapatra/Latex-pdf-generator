from email.message import EmailMessage
from credential import EMAIL_ADDRESS, EMAIL_PASSWORD
from email.message import EmailMessage
import smtplib
import firebase_admin
from firebase_admin import credentials
from firebase_admin import firestore

cred = credentials.Certificate("keys.json")
firebase_admin.initialize_app(cred)
 
db = firestore.client()

docs = (db.collection_group(u'students').get())
for doc in docs:
    email01 = doc.to_dict()['email']
    name01 = doc.to_dict()['firstName']
    SAP = doc.to_dict()['sapID']
    grad_year=doc.to_dict()['graduationYear']
    course=doc.to_dict()['course']


    msg = EmailMessage()
    msg['Subject'] = 'Generated CV'
    msg['From'] = EMAIL_ADDRESS
    msg['Cc'] = ['','']
    msg['Bcc'] = ['']
    # msg['To'] = email01  
    msg['To'] = "ritish20mohapatra@gmail.com"
    msg.set_content('Your CV is Here:') 

    msg.add_alternative("""\
		<!DOCTYPE html>
		<html>
			<body>
				<h3> Your Message </h3>
			</body>
		</html>    
		""", subtype='html')
    try:
        with open(f'pdf/{course}-{grad_year}/{SAP}.pdf', 'rb') as f:
            file_data = f.read()
            file_name = f.name
        
        with open(f'pdf/{course}-{grad_year}/{SAP}.tex', 'rb') as tex01:
            tex_data = r''''''
            tex_data = tex01.read().decode("utf-8")


            tex_name = tex01.name
            print(tex_data)

        with open(f'templates/logo.jpeg', 'rb') as img01:
            img_data = img01.read()
            img_name = img01.name

        msg.add_attachment(file_data, maintype='application',
                        subtype='octet-stream', filename=f"{SAP}.pdf")

        msg.add_attachement()
        msg.add_attachement(tex_data, maintype='application',
                        subtype='octet-stream', filename=f"{SAP}.tex")

        msg.add_attachment(img_data, maintype='application',
                        subtype='octet-stream', filename=img_name)

        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            smtp.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
            try:
                smtp.send_message(msg)
            except Exception as er:
                print(er)
    except Exception as er2:
        print(er2)