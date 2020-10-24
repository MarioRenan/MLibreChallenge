import xlrd
import smtplib
from conf.settings import gmail_user, gmail_password

wb = xlrd.open_workbook('c:\\Challenge Compliance\\files\\db.xls')
sheet = wb.sheet_by_index(0)


for i in range(1, sheet.nrows):
	if sheet.cell_value(i,4) == 'high':
		print(sheet.cell_value(i,0))
		sent_from = gmail_user
		to = [f'{sheet.cell_value(i,6)}']
		subject = 'Informação Sensível'
		body = f'\nOlá! \nEncontramos informações classificadas como altas na dn {sheet.cell_value(i,0)} \nSolicitação feita por {sheet.cell_value(i,2)}! \nData: {sheet.cell_value(i,5)} \nPoderia aprovar a classificação por favor?'

		email_text = """\
		From: %s
		To: %s
		Subject: %s

		%s
		""" % (sent_from, ", ".join(to), subject, body)

		#try:
		server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
		server.ehlo()
		server.login(gmail_user, gmail_password)
		server.sendmail(sent_from, to, email_text.encode('utf-8'))
		server.close()

		print('E-mail enviado com sucesso!')
		#except Exception as e:
			#print(e)



