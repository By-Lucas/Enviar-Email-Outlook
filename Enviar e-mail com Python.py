#Enviar email usando Outlook

#Acompanhe mais projetos no Github
# https://github.com/By-Lucas


import win32com.client as win32

# criar a integração com o outlook
outlook = win32.Dispatch('outlook.application')

# criar um email
email = outlook.CreateItem(0)

faturamento = 1500
qtde_produtos = 10
ticket_medio = faturamento / qtde_produtos

# configurar as informações do seu e-mail (PODE ADICIONAR MAIS DE 1 E-MAIL)
email.To = "tekertudo@gmail.com; pythonimpressionador+lira@gmail.com"
#Titulo do E-mail
email.Subject = "E-mail automático do Python"  
#Mensagem a ser enviada no formato HTML
email.HTMLBody = f"""  
<p>Olá Lira, aqui é o código Python</p>

<p>O faturamento da loja foi de R${faturamento}</p>
<p>Vendemos {qtde_produtos} produtos</p>
<p>O ticket Médio foi de R${ticket_medio}</p>

<p>Abs,</p>
<p>Código Python</p>
"""

#SE FOR ENVIAR ANEXO, SÓ DESCOMENTAR O CODIGO ABAIXO
# anexo = "C://Users/joaop/Downloads/arquivo.xlsx"
# email.Attachments.Add(anexo)

email.Send()
print("Email Enviado")


