import pandas as pd
import smtplib
from  email.message import EmailMessage
def enviar_email():
    clientes = pd.read_excel('./clientes.xlsx')
    for index,cliente in clientes.iterrows():
        corpo_email = f"""
        <p> <b> Prezado {cliente ['nome']}, estou testando envio de e-mail via Python,
        desconsiderar. Criei um arquivo em excel com varios nomes e e-mail para testar.
        Agora deu certo - Ok. <b> </P>
        <p> Attenciosamente, o Gerente </P>
        """
        # Compor estrutura do e-mail
        msg = EmailMessage()
        msg['Subject'] = "Assunto do email"
        msg['From']    = 'pgsguedes@gmail.com'
        msg['To']      = cliente['email']
        password = 'vcwfjljmnrwzoqxo'
        msg.add_header('Content-Type', 'text/html')
        msg.set_payload(corpo_email)
        # anexar arquivos ou imagem
        with open('logo.png', 'rb') as fp:
            img_data = fp.read()
        #    msg.add_attachment(filename='logo.png')
            msg.add_attachment(img_data, maintype='image', subtype='png',
                               filename='logo.png')
        # anexar arquivos pdf
        with open('arquivo.pdf', 'rb') as fp:
             img_data = fp.read()
        #      msg.add_attachment(filename='arquivo.pdf')
             msg.add_attachment(img_data, maintype='Application', subtype='pdf',
                                filename='arquivo.pdf')
        # anexar arquivos pdf
        with open('email_lira_2024.py', 'rb') as fp:
             img_data = fp.read()
        #     msg.add_attachment(filename='email_lira.py')
             msg.add_attachment(img_data, maintype='Application', subtype='py',
                                filename='email_lira_2024.py')
        # Servidor SMTP do Gmail
        s = smtplib.SMTP('smtp.gmail.com: 587')
        s.starttls()
        # Login credencial e enviar e-mail
        s.login(msg['From'], password)
        s.sendmail(msg['From'], [msg['To']], msg.as_string().encode('utf-8'))
        print('Email Enviado',{cliente ['nome']})
enviar_email()

