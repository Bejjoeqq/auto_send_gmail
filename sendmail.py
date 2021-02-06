import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import pandas,os

df1=pandas.read_excel(os.path.join("data.xlsx"),engine='openpyxl')

gmail_user = 'your@gmail.com'
gmail_password = 'YourPassw0rd'

server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
server.ehlo()
server.login(gmail_user, gmail_password)

subject = 'Surat Undangan Lomba Dinacom - Universitas Dian Nuswantoro Semarang'

for x in range(len(df1["Nama"])):
	msg = MIMEMultipart()
	msg['From'] = gmail_user
	msg['Subject'] = subject
	names = df1["Nama"][x]
	print(names)
	body = f'''Yth. Kepala Sekolah
{names}
di tempat

Kami panitia lomba Dinus Application Competition (DINACOM 2021) hendak mengirimkan surat undangan lomba untuk siswa/i {names} agar dapat ikut serta dalam acara yang kami selenggarakan. Bersama ini saya lampirkan surat undangan lomba. Dimohon kesediaannya untuk ikut berpartisipasi dalam acara ini. Demikian surat undangan ini saya sampaikan, atas perhatian bapak/ibu, saya ucapkan terima kasih. 

Hormat kami, 
Panitia Dinacom
Universitas Dian Nuswantoro Semarang
'''
	msg.attach(MIMEText(body, 'plain'))
	msg['To'] = df1["Email"][x]

	binary_pdf = open(f"pdf/{names}.pdf", 'rb')
	payload = MIMEBase('application', 'octate-stream', Name=f"{names}.pdf")
	payload.set_payload((binary_pdf).read())
	encoders.encode_base64(payload)
	payload.add_header('Content-Decomposition', 'attachment', filename=f"{names}.pdf")
	msg.attach(payload)

	server.sendmail(gmail_user, df1["Email"][x], msg.as_string())

server.close()