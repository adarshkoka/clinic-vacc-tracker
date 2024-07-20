import smtplib, ssl
import pandas as pd
from datetime import datetime

subject = "Default Subject Child Vacc"
message = "Default Content Child Vacc"
senderEmail = "koneello@gmail.com"
numDaysLeftToEmail = 5

class Mail:

    def __init__(self):
        self.port = 465
        self.smtp_server_domain_name = "smtp.gmail.com"
        self.sender_mail = senderEmail
        self.password = "nmqjlslccdfpmnid"

    def send(self, emails, subject, message):
        ssl_context = ssl.create_default_context()
        service = smtplib.SMTP_SSL(self.smtp_server_domain_name, self.port, context=ssl_context)
        service.login(self.sender_mail, self.password)
        
        for email in emails:
            result = service.sendmail(self.sender_mail, email, f"Subject: {subject}\n{message}")

        service.quit()

if __name__ == '__main__':

    workbook = pd.read_excel('child_vacc.xlsx')

    childNames = workbook['Child Name'].tolist()
    emails = workbook['Email Address'].tolist()
    firstDoseDates = workbook['Date of First Dose'].tolist()
    secondDoseDates = workbook['Date of Second Dose'].tolist()

    emailLogFile = open("email.log", "a")

    for i in range(len(childNames)):

        patientFirstDoseDate = firstDoseDates[i]
        patientSecondDoseDate = secondDoseDates[i]

        todaysDate = datetime.today()
        datetime_firstDose = datetime.strptime(datetime.strftime(patientFirstDoseDate, "%m/%d/%Y"), "%m/%d/%Y")
        datetime_secondDose = datetime.strptime(datetime.strftime(patientSecondDoseDate, "%m/%d/%Y"), "%m/%d/%Y")
        
        # print(patientName)
        # print(patientEmail)
        # print(datetime_firstDose)
        # print(datetime_secondDose)
        # print(todaysDate)
        # print((datetime_firstDose - datetime_secondDose).days)
        # print((datetime_secondDose - todaysDate).days)

        notifyDaysLeft = (datetime_secondDose - todaysDate).days
        # print(notifyDaysLeft)
        if notifyDaysLeft <= numDaysLeftToEmail and notifyDaysLeft > 0:

            patientName = childNames[i]
            patientEmail = [emails[i]]

            subject = "REMINDER: COVID-19 Vaccination Second Dose for your child " + patientName
            str1 = str(patientSecondDoseDate)
            str2 = str(numDaysLeftToEmail)
            # message = "Your child must receive a second dose of the COVID-19 Vaccine by {}".format(datetime_secondDose)
            # message = "Your child must receive a second dose of the COVID-19 Vaccine by " + str1 + " which is in " + str2 + " days." + "\nPlease schedule an appointment with Penguin Pediatrics."
            message = ""

            mail = Mail()
            mail.send(patientEmail, subject, message)
            logMsg = "Email sent to " + str(patientName) + " on " + str(todaysDate) + " for a Second Dose on " + str(patientSecondDoseDate) + "\n"
            emailLogFile.write(logMsg)

    emailLogFile.close()

