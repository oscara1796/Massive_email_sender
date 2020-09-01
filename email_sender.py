import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import time
import xlrd
import os




def Email_Sender(to_email, file, subject, person, text_input):
    email_user = ""
    email_send = to_email;
    Subject = subject

    msg = MIMEMultipart()
    msg['From'] = email_user
    msg['To'] = email_send
    msg['Subject'] = Subject

    html = """
    <html >
        <head>
          <meta charset="utf-8">
          <title></title>
          <style media="screen">
              table{
                margin: 0 auto

              }
              table{
                width: 600px
              }
              #logo img{
                  width: 30%%
              }
              #portada {
                width: 600px
              }
              #portada img{
                  width: 100%%

              }
              .body-text{
                  font-size: 20px
              }

              #footer-text img{
                  width: 20%%
              }

          </style>
        </head>
        <body>

          <table>
            <tr>
              <td id="logo"><img src="https://fpfjwi.stripocdn.email/content/guids/CABINET_660bbfbfdecbec035ec53be4c2508661/images/36801576606045310.png" alt=""></td>
            </tr>

          </table>
          <table>
            <tr>
              <td id="portada"> <img src="https://i.imgur.com/p8QWZxm.jpg" alt=""></td>
            </tr>
          </table>

          <table>
            <tr>
              <th><h1>¡Prepara tu negocio!  %s </h1></th>
            </tr>
            <tr>
              <td>
                <p class="body-text">Aprovecho este medio para enviarle un cordial saludo y a su vez  <br> mostrarle los productos que manejamos para preparar los negocios en esta situación actual!!
                <br>
                <br>
                A continuación encontrará anexa una cotización que conforman nuestra  cartera de productos para preparar su negocio!!
                </p>
                <p class="body-text"> %s </p>
                <p class="body-text">Quedo a su disposición para atender todas sus dudas y comentarios.</p>
                <p class="body-text">OSCAR CARRILLO</p>
              </td>
            </tr>
            <tr>
              <td id="footer-text"><img src="https://fpfjwi.stripocdn.email/content/guids/CABINET_660bbfbfdecbec035ec53be4c2508661/images/36801576606045310.png" alt=""></td>
            </tr>

          </table>

        </body>
      </html>
    """ % (person, text_input,);

    # msg.attach(MIMEText(body, 'plain'))
    msg.attach(MIMEText(html, 'html'))


    if file:
        filename = file
        attachment = open('..\..\..\..\..\Cotizaciones\Cotizaciones 2020\{}'.format(filename), 'rb')

        part = MIMEBase('application','octet-stream')
        part.set_payload((attachment).read())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition","attachment; filename= "+ filename)

        msg.attach(part)

    text = msg.as_string()

    server = smtplib.SMTP('smtp.gmail.com', 587)

    server.starttls();
    # HERE the second parameter is the email password
    server.login(email_user, '')



    server.sendmail(email_user, email_send, text)
    server.quit()
    print("enviado")




if __name__ == '__main__':
    control = True;
    while control:
        os.system('cls||clear')
        menu_ans = int(input("""
        [1] Mandar un email
        [2] Mandar varios email
        [3] Salir """))


        if menu_ans == 1:

            to_email = input("A quien le vas a mandar el email ")

            subject = input(" titulo del email ")

            person = input(" A quien va dirigido ")

            ans_text =input(" vas a incluir un mensaje (si)(no)")
            ans =input(" vas a incluir un achivo (si)(no)")

            if ans.lower() == "si":
                file= input("Cual es el nombre del archivo ")
            else:
                file= ""

            if ans_text.lower() == "si":
                text_input = input("Ingresa tu mensaje ")
            else:
                text_input=""

            try:
                Email_Sender(to_email, file, subject, person, text_input)
            except Exception:
                print("No se pudo mandar el email el correo no es valido ")
            input()

        elif menu_ans == 2:

            list_email= []
            excel_name = input("Dame el nombre del archivo ")

            loc = ("{}.xlsx".format(excel_name))

            wb = xlrd.open_workbook(loc)
            sheet = wb.sheet_by_index(0)

            n_rows= sheet.nrows -763

            for row_index in range(764,n_rows):
                cell = sheet.cell(row_index,3).value
                if cell == '':
                    break;
                else:
                    if sheet.cell(row_index,4).value != '':
                        dic = {}
                        dic['Name'] = sheet.cell(row_index,3).value;
                        dic['email'] = sheet.cell(row_index,4).value;
                        list_email.append(dic)

            subject = input(" titulo del email ")
            ans_text =input(" vas a incluir un mensaje (si)(no)")
            if ans_text.lower() == "si":
                text_input = input("Ingresa tu mensaje ")
            else:
                text_input=""
            ans =input(" vas a incluir un achivo (si)(no)")
            if ans.lower() == "si":
                file= input("Cual es el nombre del archivo ")
            else:
                file= ""

            count = 1;
            hour=0
            multiof15 = [n for n in range(1,121) if n % 15 == 0]
            for email in list_email:
                try:
                    if count in multiof15:
                        Email_Sender(email['email'],file, subject, email['Name'], text_input )
                        hour+=1;
                        time.sleep(3600)
                    elif hour == 8:
                        break;
                    else:
                        Email_Sender(email['email'],file, subject, email['Name'], text_input )
                except Exception:
                    print("No se pudo enviar el email")
                count +=1
            input()

        elif menu_ans == 3:
            control = False
