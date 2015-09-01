#!/usr/bin/python

import sqlite3 as sqlite
import os
import sys
import smtplib
from datetime import date
from datetime import timedelta
from time import tzname
from time import daylight
import email
import email.mime.application

from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate
from reportlab.lib.pagesizes import letter
from reportlab.lib.pagesizes import landscape
from reportlab.lib.units import inch
from reportlab.platypus import Paragraph
from reportlab.platypus import Table
from reportlab.platypus import TableStyle
from reportlab.lib.styles import getSampleStyleSheet
import config


# Page setup
margin = .2 * inch
baseReport = 'FlightsReport.pdf'
page_font = 'Courier'
page_font_size = int(7)
rpts = ['all', 'poi', 'chk']

# Define title attributes
styles = getSampleStyleSheet()
styleHeading = styles['Heading1']
styleHeading.alignment = 1
styleHeading.backColor = colors.blue
styleHeading.textColor = colors.white
styleHeading.fontSize = 18

# Define system info
platform = sys.platform
runtime_dir = os.path.dirname(sys.argv[0])
abs_path = os.path.abspath(runtime_dir)
tz = tzname[daylight]

# Define report dates
rptDate = ((date.today() - timedelta(1)).strftime('%A  %B %d, %Y ')) + tz
sqlDate = ((date.today() - timedelta(1)).strftime('%Y-%m-%d%%'))

class dbMgr(object):
    def __init__(self, db):
        self.conn = sqlite.connect(db)
        self.conn.execute('pragma foreign_keys = on')
        self.conn.commit()
        self.cur = self.conn.cursor()

    def query(self, arg):
        self.cur.execute(arg)
        self.conn.commit()
        return self.cur

    def __del__(self):
        self.conn.close()

def calcMsgCount(value):
    cnt = 0
    for i in range(56, 66):
        try:
            cnt = float(int(value[i])) + cnt
        except TypeError:
            pass
    return str(int(cnt))


def createDoc(rows, rptType):
    flightroute = dbMgr(config.flightRoute)
    doc = SimpleDocTemplate(rptType + baseReport,
                            rightMargin=margin,
                            leftMargin=margin,
                            topMargin=margin,
                            bottomMargin=margin,
                            pagesize=landscape(letter))

    # container for the 'Flowable' objects
    elements = []

    elements.append(Paragraph(rptType + " Flights seen on:" + "  " + rptDate, styleHeading))
    # elements.append(PageBreak)

    # Set Column Headers
    even_rows = []
    poi = []
    chk = []
    colWidths = [.75 * inch] + [.8 * inch] + [1 * inch] + [2.2 * inch] + [1.45 * inch] + [.65 * inch] + [
        .65 * inch] + [.65 * inch] + [.65 * inch] + [.65 * inch] + [.65 * inch]
    rowHeights = [.16 * inch] * ((len(rows) * 2) + 2)
    index = 0
    data = [["Start Time", "Mode S", "Call Sign", "Country", "Manufacturer", "First Sqwk", "First Alt", "First GS",
             "First VR", "First Track", "#MsgRcvd"],
            ['End Time', 'Registration', 'Route', 'Owner', 'Type', 'Last Sqwk', "Last Alt", "Last GS", "Last VR",
             "Last Track"]]

    for row in rows:
        # Change "N" to "United States"
        if row[4] == 'N':
            country = 'United States'
        else:
            country = row[4]

        # Check for missing start_time
        try:
            start_time = str(row[53][11:])
        except TypeError:
            start_time = 'UNKNOWN'

        # Check for missing end_time
        try:
            end_time = str(row[54][11:])
        except TypeError:
            end_time = 'UNKNOWN'

        try:
            mode_s = str(row[3])
        except TypeError:
            mode_s = "UNKNOWN"

        try:
            registration = str(row[6]).strip()
        except TypeError:
            registration = '------'

        try:
            manufacturer = str(row[12]).strip()
        except TypeError:
            manufacturer = '------'

        try:
            mfr_type = str(row[14]).strip()
        except TypeError:
            mfr_type = '------'

        try:
            owner = str(row[21]).strip()
        except TypeError:
            owner = '-----'

        if str(row[81]).strip() == '':
            firstSquawk = '----'
        else:
            firstSquawk = str(row[81])

        if str(row[82]).strip() == '':
            lastSquawk = '----'
        else:
            lastSquawk = str(row[82])

        if str(row[75]).strip() == '':
            firstAlt = '------'
        else:
            firstAlt = str(row[75])

        if str(row[76]).strip() == '':
            lastAlt = '------'
        else:
            lastAlt = str(row[76])

        if str(row[55]).strip() == 'None':
            callsign = '------'
            route = '------'
        else:
            callsign = str(row[55]).strip()
            route = flightroute.query("select route from FlightRoute where FlightRoute.flight like '" + callsign + "'").fetchall()
            if len(route) > 0:
                route = str(route[0][0])
            else:
                route = '------'
        try:
            firstGS = str(int(float(row[73])))
        except TypeError:
            firstGS = '------'

        try:
            lastGS = str(int(float(row[74])))
        except TypeError:
            lastGS = '------'

        msg_rcvd = calcMsgCount(row)

        try:
            firstVR = str(int(float(row[77])))
        except TypeError:
            firstVR = "------"

        try:
            lastVR = str(int(float(row[78])))
        except TypeError:
            lastVR = "------"

        try:
            firstTrk = str(int(float(row[79])))
        except TypeError:
            firstTrk = '----'

        try:
            lastTrk = str(int(float(row[80])))
        except TypeError:
            lastTrk = '----'

        data.append(
            [start_time, mode_s, callsign, country, manufacturer, firstSquawk, firstAlt, firstGS, firstVR, firstTrk,
             msg_rcvd])
        data.append([end_time, registration, route, owner, mfr_type, lastSquawk, lastAlt, lastGS, lastVR, lastTrk])

        index += 1

        if index % 2 == 0:
            even_rows.append(index)

        if row[28] == 1:
            poi.append(index)

        if str(row[6]).strip() == 'None':
            chk.append(index)

    # Reminder: (column, row) starting at 0
    # (0,0) is upper left, (-1,-1) is lower right (row,0),(row,-1) is entire row

    # t = Table(data, colWidths=colWidths, rowHeights=rowHeights, repeatRows=2)
    t = Table(data, colWidths=colWidths, rowHeights=rowHeights, repeatRows=2, hAlign='LEFT')
    t.hAlign = 'LEFT'
    t.vAlign = 'CENTER'

    # Define default table attributes 
    tblStyle = TableStyle([])
    tblStyle.add('FONTSIZE', (0, 0), (-1, -1), page_font_size)
    tblStyle.add('TEXTFONT', (0, 0), (-1, -1), page_font)
    tblStyle.add('TEXTCOLOR', (0, 0), (-1, -1), colors.black)
    tblStyle.add('TOPPADDING', (0, 0), (-1, -1), 6)
    tblStyle.add('BOTTOMPADDING', (0, 0), (-1, -1), 0)
    tblStyle.add('BOX', (0, 0), (-1, 1), .25, colors.black)
    tblStyle.add('BOX', (0, 0), (-1, -1), 2, colors.black)
    tblStyle.add('INNERGRID', (0, 0), (-1, 1), 0.15, colors.gray)
    tblStyle.add('BACKGROUND', (0, 0), (-1, 1), colors.lightblue)
    tblStyle.add('BACKGROUND', (0, 2), (-1, -1), colors.white)

    for row in even_rows:
        row1 = (row) * 2
        row2 = row1 + 1
        tblStyle.add('BACKGROUND', (0, row1), (-1, row2), colors.lightgreen)

    if rptType == 'all':
        for row in poi:
            row1 = row * 2
            row2 = row1 + 1
            tblStyle.add('BACKGROUND', (0, row1), (-1, row1), colors.red)
            tblStyle.add('BACKGROUND', (0, row2), (-1, row2), colors.red)
            tblStyle.add('TEXTCOLOR', (0, row1), (-1, row1), colors.white)
            tblStyle.add('TEXTCOLOR', (0, row2), (-1, row2), colors.white)

        for row in chk:
            row1 = row * 2
            row2 = row1 + 1
            tblStyle.add('BACKGROUND', (0, row1), (-1, row1), colors.yellow)
            tblStyle.add('BACKGROUND', (0, row2), (-1, row2), colors.yellow)
            tblStyle.add('TEXTCOLOR', (0, row1), (-1, row1), colors.black)
            tblStyle.add('TEXTCOLOR', (0, row2), (-1, row2), colors.black)

    t.setStyle(tblStyle)
    elements.append(t)

    # write the document to disk
    doc.build(elements)


def dbExtract(db, rptType):
    basestation = dbMgr(db)
    rows = []
    if rptType == 'all':
        SQL = "select * from Aircraft INNER JOIN Flights ON (Aircraft.AircraftID=Flights.AircraftID) where (Flights.EndTime like '" + sqlDate + "'" + " OR Flights.StartTime like '" + sqlDate + "'" + ") ORDER BY Flights.StartTime"
    elif rptType == 'poi':
        SQL = "select * from Aircraft INNER JOIN Flights ON (Aircraft.AircraftID=Flights.AircraftID) where (Flights.EndTime like '" + sqlDate + "'" + " OR Flights.StartTime like '" + sqlDate + "'" + ")  and Aircraft.Interested = '1' ORDER BY Flights.StartTime"
    elif rptType == 'chk':
        SQL = "select * from Aircraft INNER JOIN Flights ON (Aircraft.AircraftID=Flights.AircraftID) where  (Flights.EndTime like '" + sqlDate + "'" + " OR Flights.StartTime like '" + sqlDate + "'" + ") and Aircraft.Registration is NULL ORDER BY Flights.StartTime"

    rows = basestation.query(SQL).fetchall()
        
    return(rows, rptType)

# Extract flight info for the day

for rpt in rpts:
    rptType = rpt

    rows, rptType = dbExtract(config.baseStation, rptType)
    os.nice(15)
    createDoc(rows, rptType)

if config.sendMail:
    msg = email.mime.Multipart.MIMEMultipart()
    msg['From'] = config.sender
    msg['To'] = ", ".join(config.recipients)
    msg['Subject'] = "Flight Report for " + rptDate

    # body = email.mime.text.MIMEText("Plane report")
    body = email.mime.Text.MIMEText("")
    msg.attach(body)
    for rpt in rpts:
        fp = open(rpt + baseReport, 'rb')
        att = email.mime.application.MIMEApplication(fp.read(), _subtype="pdf")
        fp.close()
        att.add_header('Content-Disposition', 'attachment', filename=rpt + baseReport)
        msg.attach(att)

    s = smtplib.SMTP(config.smtpserver)
    if config.smtpAuth:
        s.starttls()
        s.login(config.smtpAuth) 
    s.sendmail(config.sender, config.recipients, msg.as_string())
    s.quit()
