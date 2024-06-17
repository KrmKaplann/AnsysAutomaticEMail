import sys
import clr

clr.AddReference('System')
import System.Net as net
import System.Net.Mail as smtp
import json
import time


def DownloadFile(Object, files):
    if Object.GetType().Name != "TreeGroupingFolder":
        Object.Activate()
        ImagineName = Object.Name
        image_settings = Ansys.Mechanical.Graphics.GraphicsImageExportSettings()
        image_settings.CurrentGraphicsDisplay = False
        image_settings.Resolution = GraphicsResolutionType.NormalResolution
        image_settings.Capture = GraphicsCaptureType.ImageAndLegend
        image_settings.Background = GraphicsBackgroundType.White
        image_settings.FontMagnification = 1

        ImagesLink = "C:\Users\krmka\OneDrive\Ansys Studies\SendMail\Images\\"
        ImagesFileName = ImagineName + ".png"

        Graphics.ExportImage(ImagesLink + ImagesFileName, GraphicsImageExportFormat.PNG)
        files.append(ImagesLink + ImagesFileName)


# JSON dosyasından e-posta yapılandırma bilgilerini oku
with open('C:\Users\krmka\OneDrive\Ansys Studies\SendMail\config.json', 'r') as f:
    config = json.load(f)

sender_mail = config["sender_mail"]
sender_pwd = config["sender_pwd"]
recipient_mail = config["recipient_mail"]  # Burayı geçerli bir adresle değiştirin
# Images = config["Images"]

Model.Analyses[0].Solution.Solve()

for i in range(5):
    Status = Model.Analyses[0].Solution.Status
    if str(Status) == "Done":
        print("Done")
        break
    elif str(Status) == "SolveRequired":
        print("Solve Required")
        break

files = []
Imagecounts = []
TotalImages = 1
for SolutionChildren in Model.Analyses[0].Solution.Children:
    for index, SolutionChildren2 in enumerate(SolutionChildren.Children):
        if str(SolutionChildren2.GetType().Name) == "Figure":
            print("[ " + str(TotalImages) + " ]" + " " + SolutionChildren2.Name)
            Imagecounts.append("[ " + str(TotalImages) + " ]" + " " + SolutionChildren2.Name)
            TotalImages += 1
            DownloadFile(SolutionChildren2, files)

TotalTime = Model.Analyses[0].Solution.ElapsedTime
TotalResultFileSize = int(Model.Analyses[0].Solution.ResultFileSize) / int(1000)

list_items = "".join("<li>%s</li>" % item for item in Imagecounts)

# MailMessage nesnesini oluştur
msg = smtp.MailMessage()
msg.From = smtp.MailAddress(sender_mail)
msg.To.Add(smtp.MailAddress(recipient_mail))
msg.Subject = "Ansys Mechanical Status Update"

# Creating the HTML body with the list embedded using % formatting
htmlBody = """<!DOCTYPE html>
<html>
<head>
    <title>Analysis Report</title>
</head>
<body>
    <p>The analysis has been successfully completed. The total computation time for the performed analysis is approximately <strong>%s</strong> seconds. The total size on disk is around <strong>%s</strong> KB.</p>
    <p>You can find the images you added in the attachment to view the results of the analysis.</p>
    <ul>
        %s
    </ul>
</body>
</html>""" % (TotalTime, TotalResultFileSize, list_items)

msg.Body = htmlBody
msg.IsBodyHtml = True

for file in files:
    print(file)
    attachment = smtp.Attachment(file)
    msg.Attachments.Add(attachment)

# SMTP sunucusuna giriş yap ve e-postayı gönder
client = smtp.SmtpClient('smtp-mail.outlook.com', 587)
client.EnableSsl = True
client.Credentials = net.NetworkCredential(sender_mail, sender_pwd)
client.Send(msg)
