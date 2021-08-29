import subprocess
import smtplib
import os
from email.message import EmailMessage
import shutil
import pandas as pd
from credential import EMAIL_ADDRESS, EMAIL_PASSWORD
import firebase_admin
from firebase_admin import credentials
from firebase_admin import firestore
import fire

cred = credentials.Certificate("keys.json")
firebase_admin.initialize_app(cred)

db = firestore.client()


def strFormatter(text):
	for x in range(len(text)):
		if x == 0 and text[x] == "&":
			text = "\\" + text
		elif text[x] == "&" and text[x-1] != "\\":
			text = text[:x] + "\\" + text[x:]
		elif x == 0 and text[x] == "%":
			text = "\\" + text
		elif text[x] == "%" and text[x-1] != "\\":
			text = text[:x] + "\\" + text[x:]
		elif x == 0 and text[x] == "#":
			text = "\\" + text
		elif text[x] == "#" and text[x-1] != "\\":
			text = text[:x] + "\\" + text[x:]
	encoded_string = text.encode("ascii", "ignore")
	decode_string = encoded_string.decode()
	return text
# Generating PDF
def make_pdf(s_dict):

    # Read all the templates in memory
    main_template = r''''''		# Append all sub_templates to this at last stage
    internship_template = r''''''  # Append concatenated interships to this template
    internship_parent = r''''''  # Append concatenated projects to parent template
    project_template = r''''''		# Append concatenated projects to this template
    project_parent = r''''''		# Append concatenated projects to parent template
    extracur_template = r''''''		# Append concatenated projects to extracu template
    leadership_template = r''''''		# Append concatenated leaderships to this template
    leadership_parent = r''''''		# Append concatenated projects to parent template
    leadership_parent = r''''''		# Append concatenated projects to parent template 
    git_opt = r''''''		# Github Optional
    ending = r''''''		# Ending

    with open('templates/ending.tex', 'r') as end:
            ending_template=end.read()


    # Set filename
    filename = s_dict['sapID']

# Adding github url as optional parameter
    git_lenght=len(s_dict["githubUrl"])
    if git_lenght==0:
        git_template = r''''''
        git_template=" "
    else:
        with open('templates/git_template.tex', 'r') as gits:
            git_opt=gits.read()
            git_template = r''''''
            git_template=git_opt %(("githubUrl"),"GitHub")
            

    # Interpolate data into internship_template
    internship_lenght=len(s_dict["internship"])
    if internship_lenght==0:
        pass
    else:
        with open('templates/internship_parent.tex', 'r') as fg:
            internship_parent=fg.read()
        for internship in s_dict["internship"]:
            
            with open('templates/internship_template.tex', 'r') as f:

                intern_temp = r''''''
                intern_temp = f.read()

                description_main = r''''''

                for description in internship["internDesc"]:

                    with open('templates/item_template.tex', 'r') as p:

                        temp = r''''''
                        temp = p.read()

                        description_main += temp % (strFormatter(description))

                internship_template += intern_temp % (
                    strFormatter(internship["orgName"]), strFormatter(internship["internDur"]), strFormatter(internship["internRole"]), description_main)
        internship_parent =internship_parent%(internship_template)


# Interpolate data into project_template
    project_lenght=len(s_dict["project"])
    if project_lenght==0:
        pass
    else:
        with open('templates/project_parent.tex', 'r') as fpr:
            project_parent=fpr.read()
        for count, project in enumerate(s_dict["project"]):
            
            with open('templates/project_template.tex', 'r') as f:

                proj_temp = r''''''
                proj_temp = f.read()

                description_main = r''''''

                for description in project["projDesc"]:

                    with open('templates/item_template.tex', 'r') as p:

                        temp = r''''''
                        temp = p.read()

                        description_main += temp % (strFormatter(description))

                project_template += proj_temp % (
                    strFormatter(project["projName"]), strFormatter(project["projDur"]), strFormatter(project["projTool"]), description_main)
        project_parent=project_parent%(project_template)



# Interpolate data into Extracurricular_Templates
    
    with open('templates/extracur.tex', 'r') as extra:
        extracur_template=extra.read()    
        extracur_temp = r''''''
 
        extracur_first=strFormatter(s_dict['hobbies'])
        extracur_second=strFormatter(s_dict['certificationAndCourse'])
        extracur_template =  extracur_template % (extracur_first,extracur_second)

# courses implementation 
    if s_dict['course']=="MBATech":
        s_dict['course']="BTech(Computer science), MBA"


# Interpolate data into leadership_template
    leadership_lenght=len(s_dict["leadership"])
    if leadership_lenght==0:
        pass   
    else:
        with open('templates/leadership_parent.tex', 'r') as fl:
            leadership_parent=fl.read()
        for leadership in s_dict["leadership"]:                     

            with open('templates/leadership_template.tex', 'r') as f:

                lead_temp = r''''''
                lead_temp = f.read()

                description_main = r''''''

                for description in leadership["leadDesc"]:

                    with open('templates/item_template.tex', 'r') as p:

                        temp = r''''''
                        temp = p.read()

                        description_main += temp % (strFormatter(description))

                leadership_template += lead_temp % (strFormatter(leadership["leadName"]), strFormatter(leadership["leadDur"]),
                                                    strFormatter(leadership["leadRole"]), description_main)
        leadership_parent=leadership_parent%(leadership_template)    





    # Read main template from tex file
    with open('templates/resume_template.tex', 'r') as f:
        main_template = f.read() 

    # Interpolate data into main_template
    dat = main_template % (s_dict['firstName'],
                           s_dict['lastName'],
                           s_dict['mobile'],
                           s_dict['mobile'],
                           strFormatter(s_dict['linkedinUrl']),
                           git_template,
                           s_dict['email'],
                           s_dict['email'],
                           s_dict['admissionYear'],
                           s_dict['graduationYear'],
                           s_dict['course'],                        
                           s_dict['cgpa'],
                           strFormatter(s_dict['programmingLanguage']),
                           strFormatter(s_dict['toolsAndTechnologies']),
                           strFormatter(s_dict['coreSkills'])                                       
                           )
    if internship_lenght>0:
        dat+=internship_parent
    if project_lenght>0:
        dat+=project_parent
    dat+=extracur_template
    if leadership_lenght>0:
        dat+=leadership_parent    
    dat+=ending_template

    # Write new template file
    with open(f"{filename}.tex", 'w') as f:
        f.write(dat)
    # Calling CLI utility
    proc = subprocess.Popen(['pdflatex', f"{filename}.tex"])
    proc.communicate()

    # Redundant file removal, may need logs for debugging
    os.unlink(f"{filename}.aux")
    os.unlink(f"{filename}.log")
    os.unlink(f"{filename}.out")
    cleanUp(s_dict['sapID'], s_dict['course'], s_dict['graduationYear'])

# File organization
def cleanUp(sap, batch, gradYear):
	batchDir = f'{batch}-{gradYear}'
	if not(os.path.exists('pdf')):
		os.mkdir('pdf')
	if not(os.path.exists(os.path.join('pdf', batchDir))):
		os.mkdir(os.path.join('pdf', batchDir))
	files = [f for f in os.listdir('.') if os.path.isfile(f) and (f.endswith('.pdf') or f.endswith('.tex'))]
	for stuFile in files:
		shutil.move(stuFile, os.path.join('pdf', batchDir))


# def cleanUp():
# 	# File organization
# 	os.mkdir('pdf')
# 	files = [f for f in os.listdir('.') if os.path.isfile(f) and f.endswith('.pdf')]
# 	for pdf in files:
# 		shutil.move(pdf, 'pdf')

def createAll():
	docs = db.collection_group("students").get()
	for doc in docs:
		py_dict = doc.to_dict()
		make_pdf(py_dict)


def createBatch(batch='None'):
	docs = db.collection('student').document(f"{batch}").collection('students').get()
	for doc in docs:
		py_dict = doc.to_dict()
		make_pdf(py_dict)


def createOne(sapid='0000'):
	docs = db.collection_group('students').where('sapID', '==', str(sapid)).stream()
	for doc in docs:
		py_dict = doc.to_dict() 
		make_pdf(py_dict)


if __name__ == '__main__':
	fire.Fire()


# file organization
os.mkdir('pdf')
files = [f for f in os.listdir('.') if os.path.isfile(f) and f.endswith('.pdf')]
for pdf in files:
    shutil.move(pdf, 'pdf')


# Sending Email
# docs = (db.collection_group(u'students').get())
# for doc in docs:
#     email01 = doc.to_dict()['email']
#     name01 = doc.to_dict()['firstName']


#     msg = EmailMessage()
#     msg['Subject'] = 'CV'
#     msg['From'] = EMAIL_ADDRESS
#     # msg['To'] = email01
#     msg['To'] = "ritish20mohapatra@gmail.com"
#     msg.set_content('Your CV is Here:')

#     msg.add_alternative("""\
# 		<!DOCTYPE html>
# 		<html>
# 			<body>
# 				<h3 >Here's your generated CV.
#             Regards, 
#             Placement committee, STME NMIMS Navi Mumba</h3>
# 			</body>
# 		</html>    
# 		""", subtype='html')

#     with open(f'pdf/{name01}.pdf', 'rb') as f:
#         file_data = f.read()
#         file_name = f.name

#     msg.add_attachment(file_data, maintype='application',
#                        subtype='octet-stream', filename=file_name)

    # with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
    #     smtp.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
    #     smtp.send_message(msg)
