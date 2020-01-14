#necessary packages
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import email
import email.mime.application
import smtplib
import csv

# csv path
csv_path = 'recipients.csv' 

with open(csv_path, 'r') as csv:
    reader = csv.reader(csv, delimiter=',')
    # get header from first row
    headers = next(reader)
    # get all the rows as a list
    recipients = list(reader)
    
#recipients -> matrix: [[name_recipient, email_recipient], ...]

for recipient in recipients:
    try:
        msg = MIMEMultipart()

        #informations
        password = 'password' #Your password
        msg['From'] = 'name@exemple.com' #Your email

        #recipient
        msg['To'] = recipient[1].strip()
        
        #Subject
        msg['Subject'] = "Choose a Subject"
        
        #html
        html = str.format("""\
        <html>
          <body>
          <p> Dear {},</p> 
          <p>Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. Cum sociis natoque penatibus et magnis dis parturient montes, nascetur ridiculus mus.</p>
          <p>Donec quam felis, ultricies nec, pellentesque eu, pretium quis, sem. Nulla consequat massa quis enim. Donec pede justo, fringilla vel, aliquet nec, vulputate eget, arcu.</p>
          <p>In enim justo, rhoncus ut, imperdiet a, venenatis vitae, justo. Nullam dictum felis eu pede mollis pretium. Integer tincidunt. Cras dapibus. Vivamus elementum semper nisi. Aenean vulputate eleifend tellus. Aenean leo ligula, porttitor eu, consequat vitae, eleifend ac, enim. Aliquam lorem ante, dapibus in, viverra quis, feugiat a,</p>
          </body>
        </html>
        """, recipient[0])

        #attachments names
        files_names = ['sample.pdf', 'pptexample.ppt']

        # Adding pdf or ppt file attachment
        for file_name in files_names:
            fo = open(file_name,'rb')
            file_type = file_name.split('.')[-1]
            attach = email.mime.application.MIMEApplication(fo.read(), _subtype = file_type)
            fo.close()
            attach.add_header('Content-Disposition','attachment',filename=file_name)
            # Add an attachment to body message.
            msg.attach(attach)

        # add html
        msg.attach(MIMEText(html, 'html'))

        #create server
        server = smtplib.SMTP('smtp.gmail.com: 587') #We use gmail in this example

        server.starttls()

        # Login Credentials for sending the mail
        server.login(msg['From'], password)

        # send the message via the server.
        server.sendmail(msg['From'], msg['To'], msg.as_string())

        server.quit()

        print('The email was successfully sent to: ' + recipient[0])

    except:
        print('There was an error sending an email to: ' + recipient[0])