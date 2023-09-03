# title: "FastDealReminder"
# description: "This script reminds shift-colleagues about missing FastDeal-files and sends an email."
# output: "email"
# parameters: {}
# owner: "MCSO, Lukas Dicke"

"""

Usage:
    FastDealReminder.py <job_path> --daysAhead=<int>

Options:
    --daysAhead=<int> rel. days to today => Delivery-Day [default: 1].

"""

import json
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
import smtplib
from pathlib import Path
from datetime import datetime, timedelta

CONST_SUBSTRING_FILENAME = "SubstringFilename"
CONST_LOCATION = "Location"
CONST_DELIVERY_DAY_INDEX = "DeliveryDayIndex"
CONST_INFO_TEXT = "InfoText"
CONST_DELIVERY_DAY_FORMAT = "DeliveryDateFormat"
CONST_WARNING_TIMESTAMP = "WarningTimestamp"
CONST_CLEARNAME = "Clearname"


def SendMailPythonServer(send_to, send_cc, send_bcc, subject, body, files=[]):
    msgBody = """<html><head></head>
        <style type = "text/css">
            table, td {height: 3px; font-size: 14px; padding: 5px; border: 1px solid black;}
            td {text-align: left;}
            body {font-size: 12px;font-family:Calibri}
            h2,h3 {font-family:Calibri}
            p {font-size: 14px;font-family:Calibri}
         </style>"""

    msgBody += "<h2>" + subject + "</h2>"
    # msgBody += "<h3>" + message + "</h3>"
    msgBody += body

    strFrom = "no-reply-duswvpyt002p@statkraft.de"

    msgRoot = MIMEMultipart('related')
    msgRoot['Subject'] = subject
    msgRoot['From'] = strFrom
    if len(send_to) == 1:
        msgRoot['To'] = send_to[0]
    else:
        msgRoot['To'] = ",".join(send_to)

    if len(send_cc) == 1:
        msgRoot['Cc'] = send_cc[0]
    else:
        msgRoot['Cc'] = ",".join(send_cc)

    if len(send_cc) == 1:
        msgRoot['Bcc'] = send_bcc[0]
    else:
        msgRoot['Bcc'] = ",".join(send_bcc)
    msgRoot.preamble = 'This is a multi-part message in MIME format.'

    msgAlternative = MIMEMultipart('alternative')
    msgRoot.attach(msgAlternative)

    for path in files:
        part = MIMEBase('application', "octet-stream")
        with open(path, 'rb') as file:
            part.set_payload(file.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition',
                        'attachment; filename={}'.format(Path(path).name))
        msgRoot.attach(part)

    msgText = MIMEText('Sorry this mail requires your mail client to allow for HTML e-mails.')
    msgAlternative.attach(msgText)

    msgText = MIMEText(msgBody, 'html')
    msgAlternative.attach(msgText)

    smtp = smtplib.SMTP('smtpdus.energycorp.com')
    smtp.sendmail(strFrom, send_to, msgRoot.as_string())
    smtp.quit()

    print("Mail sent successfully from " + strFrom)


def getUncPathFastDealReminderConfig():
    return "\\\\energycorp.com\\common\\divsede\\Operations\\Exchanges\\FDUpload\\FastDealReminderConfig.json"


def getNeedOfWarning(strTimestamp):
    today = datetime.today()

    todayStartingMidnight = datetime(today.year, today.month, today.day)

    warningTimestamp = todayStartingMidnight + timedelta(hours=int(strTimestamp[0:2])) + timedelta(
        minutes=int(strTimestamp[-2:]))

    if warningTimestamp < datetime.now():
        return True
    else:
        return False


def ConvertHtmlStringToPlainText(text):
    return str.replace(text, "<br>", "\r\n")


class FastDealReminderConfig():
    def __init__(self, _config):
        self.DayIndex = int(_config[CONST_DELIVERY_DAY_INDEX])
        self.DeliveryDayFormat = _config[CONST_DELIVERY_DAY_FORMAT]
        self.ProcessFolder = _config[CONST_LOCATION]
        self.ArchiveFolder = _config[CONST_LOCATION] + "Archive\\"
        self.Substring = _config[CONST_SUBSTRING_FILENAME]
        self.Clearname = _config[CONST_CLEARNAME]
        _timedelta = timedelta(days=testdaysAhead + self.DayIndex)
        self.DeliveryDay = datetime.today() + _timedelta
        self.NeedOfWarning = getNeedOfWarning(_config[CONST_WARNING_TIMESTAMP])
        self.StrDeliveryDayToSearchFor = self.DeliveryDay.strftime(self.DeliveryDayFormat)
        self.StrDeliveryDayToReport = self.DeliveryDay.strftime("%d.%m.%y")
        self.WarningMessage = ""


def GetHyperlink(filename):
    return "<a href=" + filename + ">Click to open folder</a>"


def GetFiles(_config, folder):
    from os import listdir
    from os.path import isfile, join
    return [file for file in listdir(folder)
            if isfile(join(folder, file))
            if _config.Substring in file
            if _config.StrDeliveryDayToSearchFor in file
            ]

def CheckMissingFileInArchiveFolder(_config):

    folder=_config.ArchiveFolder

    if len(GetFiles(_config, folder)) == 0:
        _config.WarningMessage = "<br>" + "No appropriate file '" + _config.Clearname + "'  (Del.: " + _config.StrDeliveryDayToReport + ") found (" + GetHyperlink(_config.ArchiveFolder) + ")."

    return _config


def CheckThroughWholeConfig(config1):
    messageMissingFiles = ""

    checkedConfigs = []

    for configEntry in config1:

        config = FastDealReminderConfig(configEntry)

        if config.NeedOfWarning:
            config = CheckMissingFileInArchiveFolder(config)

            checkedConfigs.append(config)

    return checkedConfigs


recipientsTo = ["lukas.dicke@statkraft.de"]

testdaysAhead = 1

emailBody = ""
emailSubject = "FastDealReminderTest"

message = ""

f = open(getUncPathFastDealReminderConfig())

configs = json.load(f)

_checkedConfigs = CheckThroughWholeConfig(configs)

for config in _checkedConfigs:
    message = message + config.WarningMessage

SendMailPythonServer(send_to=recipientsTo, send_cc=[], send_bcc=[], subject=emailSubject, body=message, files=[])

print(ConvertHtmlStringToPlainText(message))
