from tkinter import *  
from tkinter import filedialog
from openpyxl.workbook import Workbook
from openpyxl import load_workbook
import pandas as pd 
import smtplib as sm 
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from email.mime.text import MIMEText


root = Tk()  
  
root.geometry("400x250")  
root.configure(bg = "#673ab7")
root.title("Bulk email sender")
# root.iconbitmap("mail.ico") 

def mailSender():
	email = userValue.get()
	password = passValue.get()
	root2 = Toplevel()
	root2.geometry("900x650")
	root2.configure(bg="#673ab7")
	root2.title("Bulk email sender")
	# root2.iconbitmap("mail.ico")


	def filePath():
		root.filename = filedialog.askopenfilename(initialdir="/",title="choose csv file",filetypes= (("Csv file","*.csv"),("All file","*.*")))
		path = root.filename
		global newPath
		newPath = ""
		k = 0
		for i in range(len(path)):

			if path[i]=='/':
				newPath = newPath+path[k:i]+'\\\\'
				k = i+1

		Label(root2,text="File Successfully loaded!",fg='green').place(x = 300,y = 20)	
		Label(root2,text = path).place(x = 10,y = 50)		
        		



	def file():

		Path = newPath + fileName.get()
		df = pd.read_csv(Path)
		global email_list 
		email_list = list(df.get("email"))
		print(email_list)
		Label(root2,text="File Successfully Fetched !",fg="green").place(x = 70, y= 120)
		

	def message():
		body = mainMessage.get("1.0", "end-1c")
		email = userValue.get()
		password = passValue.get()
		try:
			server = sm.SMTP("smtp.gmail.com", 587)
			server.starttls()
			server.login(email,password)
			print(email)
			print(password)
			from_ = email 
			to_ = email_list 
			print(to_)
			message = MIMEMultipart("alternative")
			message['Subject'] = subject.get()

			message['from'] = email 
    		 
			message.attach(MIMEText(body,'plain'))

			print(attchFile.get())
			if attchFile.get() == 'Null':
				server.sendmail(from_ , to_ , message.as_string())
			
				Label(root2,text="Message sent successfully!",fg="green").place(x = 80 ,y = 460)
				server.quit()
			else:

				Attach_filename = attchFile.get()  # In same directory as script

    		
				with open(Attach_filename, "rb") as attachment:
        			# Add file as application/octet-stream
        			# Email client can usually download this automatically as attachment
					part = MIMEBase("application", "octet-stream")
					part.set_payload(attachment.read())

    			# Encode file in ASCII characters to send by email    
				encoders.encode_base64(part)

   				# Add header as key/value pair to attachment part
				part.add_header(
					"Content-Disposition",
					f"attachment; filename= {Attach_filename}",
    			)

				message.attach(part)
				server.sendmail(from_ , to_ , message.as_string())
			
				Label(root2,text="Message sent successfully!",fg="green").place(x = 80 ,y = 600)
				server.quit()
		except Exception as E:
				print(E)

	fileName = StringVar()
	subject = StringVar() 
	attchFile = StringVar()

	Button(root2,text = "Choose File",command=filePath,bg="green",activebackground='blue').place(x = 10,y = 20)
	Label(root2,text = "Enter your filename").place(x = 10, y = 90)
	Entry(root2,textvariable = fileName).place(x = 130, y = 90)
	Button(root2,text = "Fetch file",command=file,bg="green",activebackground='blue').place(x = 10,y = 120 )
	Label(root2,text="Compose Mail").place(x=450,y=170)
	Label(root2,text = "Subject",).place(x= 10,y = 200)
	Entry(root2,textvariable = subject).place(x=80,y= 200)
	Label(root2,text = "description").place( x = 10, y = 280)
	Label(root2,text = "Enter Attach file name").place(x = 10 , y = 560)
	Entry(root2,textvariable = attchFile).place(x = 160, y = 560)
	Label(root2,text = "If no attachment present, enter --> Null").place(x = 450 , y = 560)
	mainMessage = Text(root2,width = 100,height = 15,bg = "light yellow")
	mainMessage.place(x = 80 ,y = 280)
	
	Button(root2,text = "Send",command = message,bg= "green",activebackground='blue').place(x=10, y=600)

 
  
email = Label(root, text = "Email").place(x = 50, y = 90)  
  
password = Label(root, text = "Password").place(x = 50, y = 130)  
  
sbmitbtn = Button(root, text = "Login",bg = "green", activebackground='blue',command=mailSender).place(x = 50, y = 180)  
  
userValue = StringVar()
passValue = StringVar()
  
Entry(root,textvariable=userValue).place(x = 130, y = 90)  

Entry(root,textvariable=passValue,show = "*").place(x = 130, y = 130)  

root.mainloop()  