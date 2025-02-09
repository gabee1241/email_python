# Biblioteca usada
import win32com.client as win32

# Criar a integração com o outlook
outlook= win32.Dispatch('outlook.application')

# Criar um email
email= outlook.CreateItem(0)

# Configurar as informações de seu email
email.To = "seu email!"
email.Subject= "Teste email automatico"
email.HTMLBody= """
<p>Olá Gabriel, aqui é o Python</p>

<p>Estamos fazendo um teste de automatização e enviamos este email</p>

<p>Abs, gabriel</p>
"""


email.Send()
print(f"Email enviado")

