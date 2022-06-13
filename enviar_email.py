import win32com.client as win32

#Criar a integração com o Outlook
outlook = win32.Dispatch('outlook.application')

#Criar um email
email = outlook.CreateItem(0)

#Configurar as informações do seu email
email.To = "destino"
email.Subject = "assunto"
email.HTMLBody = """
<p>corpo do email</p>
"""