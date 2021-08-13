import os
from PyQt5 import QtWidgets, QtGui
from PyQt5.QtWidgets import QFileDialog
from PyQt5.QtWidgets import QMessageBox
import pandas as pd
import re as RegEx
import datetime
import read_ini_config
from dateutil.parser import parse

# Create an application
app = QtWidgets.QApplication([])

# Create window
win = QtWidgets.QWidget()
win.setWindowTitle('CTCOB17-AuditTemplateValidationTool')
# win.setFixedSize(500, 500)
win.showMaximized()
base_path = os.getcwd()

# add application icon
icon = QtGui.QIcon()
icon.addPixmap(QtGui.QPixmap("icons/Sonic.ico"), QtGui.QIcon.Normal, QtGui.QIcon.On)
win.setWindowIcon(icon)

# set windows layout
layout = QtWidgets.QGridLayout()
win.setLayout(layout)

# read configuration file
inputPath = read_ini_config.read_config("PathConfiguration", "InputExcelFile")
logPath = read_ini_config.read_config("PathConfiguration", "LogPath")
tempPath = read_ini_config.read_config("PathConfiguration", "TempPath")
ListStatus = read_ini_config.readFile(base_path + "\\StatusList.txt")
CorporateReport = read_ini_config.read_config("PathConfiguration", "CorporatePath")
EventReport = read_ini_config.read_config("PathConfiguration", "EventPath")

# A components of widgets
# selectComboBox = QtWidgets.QComboBox()
validateButton = QtWidgets.QPushButton('Start Validation')
browseMasterPath = QtWidgets.QPushButton('Browse Shared Template')
browseCorporatePath = QtWidgets.QPushButton('Browse Corporate Data')
browseEventPath = QtWidgets.QPushButton('Browse Event Data')
masterPathText = QtWidgets.QLineEdit()
masterPathText.setReadOnly(True)
corporatePathText = QtWidgets.QLineEdit()
corporatePathText.setReadOnly(True)
eventPathText = QtWidgets.QLineEdit()
eventPathText.setReadOnly(True)
logViewer = QtWidgets.QPlainTextEdit()

# Add the QWebViews to the layout
# layout.addWidget(selectComboBox, 0, 1)
layout.addWidget(masterPathText, 0, 1)
layout.addWidget(browseMasterPath, 0, 2)
layout.addWidget(corporatePathText, 1, 1)
layout.addWidget(browseCorporatePath, 1, 2)
layout.addWidget(eventPathText, 2, 1)
layout.addWidget(browseEventPath, 2, 2)
layout.addWidget(validateButton, 3, 2)
layout.addWidget(logViewer, 4, 0, 2, 3)
font = QtGui.QFont()
font.setPointSize(12)
logViewer.setFont(font)

masterPathText.setText(inputPath)
corporatePathText.setText(CorporateReport)
eventPathText.setText(EventReport)

#read ini sheet name values
corporateSumReport_EntityVitalsSheetName = read_ini_config.read_config('CorporateSheetNames', 'SheetName1')
corporateSumReport_AuthorityToDoBusinessSheetName = read_ini_config.read_config('CorporateSheetNames', 'SheetName2')
eventSumReport_FilingEventsSheetName = read_ini_config.read_config('EventSheetNames', 'SheetName1')
masterAuditSheetName = read_ini_config.read_config('MasterSheetNames', 'SheetName1')

#read ini column name values
corporateSumReport_EntityName = read_ini_config.read_config('CorporateColumnNames', 'ColumnName1')
corporateSumReport_EntityType = read_ini_config.read_config('CorporateColumnNames', 'ColumnName2')
corporateSumReport_DomesticJurisdiction = read_ini_config.read_config('CorporateColumnNames', 'ColumnName3')
corporateSumReport_FormationDate = read_ini_config.read_config('CorporateColumnNames', 'ColumnName4')
corporateSumReport_CharterID = read_ini_config.read_config('CorporateColumnNames', 'ColumnName5')
corporateSumReport_RegisteredJurisdiction = read_ini_config.read_config('CorporateColumnNames', 'ColumnName6')
corporateSumReport_Jurisdiction = read_ini_config.read_config('CorporateColumnNames', 'ColumnName7')
corporateSumReport_RegistrationDate = read_ini_config.read_config('CorporateColumnNames', 'ColumnName8')
corporateSumReport_Status = read_ini_config.read_config('CorporateColumnNames', 'ColumnName9')

eventSumReport_EntityName = read_ini_config.read_config('EventColumnNames', 'ColumnName1')
eventSumReport_DomesticJurisdiction = read_ini_config.read_config('EventColumnNames', 'ColumnName2')
eventSumReport_Jurisdiction = read_ini_config.read_config('EventColumnNames', 'ColumnName3')
eventSumReport_DueDate = read_ini_config.read_config('EventColumnNames', 'ColumnName4')

auditReport_EntityName = read_ini_config.read_config('MasterColumnNames', 'ColumnName1')
auditReport_EntityType = read_ini_config.read_config('MasterColumnNames', 'ColumnName2')
auditReport_DomesticState = read_ini_config.read_config('MasterColumnNames', 'ColumnName3')
auditReport_FileDate = read_ini_config.read_config('MasterColumnNames', 'ColumnName4')
auditReport_StateID = read_ini_config.read_config('MasterColumnNames', 'ColumnName5')
auditReport_ForeignJurisdictions = read_ini_config.read_config('MasterColumnNames', 'ColumnName6')
auditReport_NextARDue = read_ini_config.read_config('MasterColumnNames', 'ColumnName7')
auditReport_Status = read_ini_config.read_config('MasterColumnNames', 'ColumnName8')

# check if folders in path or files exists
if not os.path.exists(inputPath):
    os.mkdir(inputPath)

if not os.path.exists(logPath):
    os.mkdir(logPath)

if not os.path.exists(tempPath):
    os.mkdir(tempPath)


def is_date(string, fuzzy=False):
    """
    Return whether the string can be interpreted as a date.

    :param string: str, string to check for date
    :param fuzzy: bool, ignore unknown tokens in string if True
    """
    try:
        parse(string, fuzzy=fuzzy)
        return True

    except ValueError:
        return False


def startValidations():
    global counterCorporateSummary, counterCorporateSummaryAuth, counterEvenSummary, corporateEntityName, corporateRegisteredJurisdiction, corporateJurisdiction, eventJurisdiction
    global eventEntityName, eventDomesticJurisdiction, eventJurisdiction, eventDueDate, found_ARDue, not_found_ARDue
    global entityname, state, jurisdicion, ARDue, filedate, stateID, corporateAuthorityRows, eventRows, event_found
    logViewer.clear()
    QtGui.QGuiApplication.processEvents()
    logViewer.appendPlainText("Reading Master Excel File...")
    QtGui.QGuiApplication.processEvents()
    masterExcel = pd.read_excel(masterPathText.text(), header=0, keep_default_na=False, sheet_name=masterAuditSheetName)
    masterExcelDrop = pd.DataFrame(masterExcel)
    masterExcelDrop = masterExcelDrop.drop([masterExcelDrop.index[0], masterExcelDrop.index[1], masterExcelDrop.index[2], masterExcelDrop.index[3], masterExcelDrop.index[4], masterExcelDrop.index[5], masterExcelDrop.index[6], masterExcelDrop.index[7], masterExcelDrop.index[8]])
    masterExcelFinal = pd.DataFrame(masterExcelDrop)
    masterExcel = masterExcelFinal.iloc[0]
    masterExcelFinal = masterExcelFinal[1:]
    masterExcelFinal.columns = masterExcel
    masterExcelFinal.drop(masterExcelFinal.columns[0], axis=1, inplace=True)
    masterExcelFinal.reset_index(drop=True, inplace=True)
    logViewer.appendPlainText("[Reading Corporate Summary Report...]")
    QtGui.QGuiApplication.processEvents()
    corporateSumReport_EntityVitals = pd.read_excel(corporatePathText.text(), header=0, keep_default_na=False, sheet_name=corporateSumReport_EntityVitalsSheetName)
    corporateSumReport_AuthorityToDoBusiness = pd.read_excel(corporatePathText.text(), header=0, keep_default_na=False, sheet_name=corporateSumReport_AuthorityToDoBusinessSheetName)

    #create dataframe
    corporateSumReport_EntityVitals = pd.DataFrame(corporateSumReport_EntityVitals)
    corporateSumReport_AuthorityToDoBusiness = pd.DataFrame(corporateSumReport_AuthorityToDoBusiness)

    # Start compare
    # compare Corporate Summary Report under Authority to do Business sheet
    for auditSheetRows in masterExcelFinal.index:
        auditSheetEntityName = masterExcelFinal[auditReport_EntityName][auditSheetRows]
        auditSheetEntityType = masterExcelFinal.iloc[auditSheetRows, 2]
        auditSheetDomesticState = masterExcelFinal[auditReport_DomesticState][auditSheetRows]
        auditSheetForeignJurisdictions = masterExcelFinal[auditReport_ForeignJurisdictions][auditSheetRows]
        auditSheetFileDate = masterExcelFinal[auditReport_FileDate][auditSheetRows]
        auditSheetStateID = masterExcelFinal[auditReport_StateID][auditSheetRows]
        if auditSheetFileDate != "":
            try:
                auditSheetFileDate = datetime.datetime.strptime(str(auditSheetFileDate), "%Y-%m-%d %H:%M:%S").strftime("%m-%d-%Y")
            except ValueError:
                logViewer.appendPlainText("Row # [" + str(auditSheetRows + 12) + "] Entity Name[" + str(auditSheetEntityName) + "] StateID[" + str(auditSheetStateID) +"] File Date[" + auditSheetFileDate + "] FileDate has problem!!!")
                QtGui.QGuiApplication.processEvents()
                break
        auditSheetStatus = masterExcelFinal[auditReport_Status][auditSheetRows]
        auditSheetNextARDue = masterExcelFinal[auditReport_NextARDue][auditSheetRows]
        corporateAuthorityRows = 0
        counterCorporateSummaryAuth = 0
        for corporateAuthorityRows in corporateSumReport_AuthorityToDoBusiness.index:
            corporateEntityName = corporateSumReport_AuthorityToDoBusiness[corporateSumReport_EntityName][corporateAuthorityRows]
            corporateRegisteredJurisdiction = corporateSumReport_AuthorityToDoBusiness[corporateSumReport_RegisteredJurisdiction][corporateAuthorityRows]
            corporateJurisdiction = corporateSumReport_AuthorityToDoBusiness[corporateSumReport_Jurisdiction][corporateAuthorityRows]
            if auditSheetForeignJurisdictions == "":
                auditSheetForeignJurisdictions = corporateJurisdiction.rstrip()
            corporateRegistrationDate = corporateSumReport_AuthorityToDoBusiness[corporateSumReport_RegistrationDate][corporateAuthorityRows]
            corporateCharterID = corporateSumReport_AuthorityToDoBusiness[corporateSumReport_CharterID][corporateAuthorityRows]
            corporateStatus = corporateSumReport_AuthorityToDoBusiness[corporateSumReport_Status][corporateAuthorityRows]
            if auditSheetFileDate == corporateRegistrationDate:
                filedate = 1
            else:
                filedate = 0
            if auditSheetStateID == corporateCharterID:
                stateID = 1
            else:
                stateID = 0

            if auditSheetEntityName == corporateEntityName and auditSheetDomesticState == corporateRegisteredJurisdiction and auditSheetForeignJurisdictions == corporateJurisdiction: # and auditSheetStateID == corporateCharterID and auditSheetStatus == corporateStatus or str(auditSheetFileDate) == str(corporateRegistrationDate):
                if filedate == 0:
                    logViewer.appendPlainText("Row # [" + str(auditSheetRows + 12) + "][" + str(auditSheetFileDate) + "] File Date in Audit Sheet cannot be found in Corporate Summary Report under Authority to do Business sheet.")
                    QtGui.QGuiApplication.processEvents()
                    filedate = 0
                if stateID == 0:
                    logViewer.appendPlainText("Row # [" + str(auditSheetRows + 12) + "][" + str(auditSheetStateID) + "] State ID # in Audit Sheet cannot be found in Corporate Summary Report under Authority to do Business sheet.")
                    QtGui.QGuiApplication.processEvents()
                    stateID = 0
                corporateAuthorityRows = 0
                counterCorporateSummaryAuth = counterCorporateSummaryAuth + 1
                break
        if counterCorporateSummaryAuth == 0:
            logViewer.appendPlainText("Row # [" + str(auditSheetRows + 12) + "] Entity Name[" + str(auditSheetEntityName) + "] Domestic State[" + str(auditSheetDomesticState) + "] Jurisdictions[" + str(auditSheetForeignJurisdictions) + "] cannot be found in Corporate Summary Report under Authority to do Business sheet.")
            QtGui.QGuiApplication.processEvents()
            counterCorporateSummaryAuth = 0

    logViewer.appendPlainText("\r\n[Reading Event Summary Report...]")
    QtGui.QGuiApplication.processEvents()
    eventSumReport_FilingEvents = pd.read_excel(eventPathText.text(), header=0, keep_default_na=False, sheet_name=eventSumReport_FilingEventsSheetName)
    eventSumReport_FilingEvents = pd.DataFrame(eventSumReport_FilingEvents)
    eventSumReport_FilingEvents2 = pd.read_excel(eventPathText.text(), header=0, keep_default_na=False, sheet_name=eventSumReport_FilingEventsSheetName)
    eventSumReport_FilingEvents2 = pd.DataFrame(eventSumReport_FilingEvents)

    # compare Event Summary Report under Filing Events sheet
    auditSheetRows = 0
    for auditSheetRows in masterExcelFinal.index:
        auditSheetEntityName = masterExcelFinal[auditReport_EntityName][auditSheetRows]
        auditSheetEntityType = masterExcelFinal.iloc[auditSheetRows, 2]
        auditSheetDomesticState = masterExcelFinal[auditReport_DomesticState][auditSheetRows]
        auditSheetForeignJurisdictions = masterExcelFinal[auditReport_ForeignJurisdictions][auditSheetRows]
        if auditSheetForeignJurisdictions == "":
            auditSheetForeignJurisdictions = auditSheetDomesticState
        auditSheetFileDate = masterExcelFinal[auditReport_FileDate][auditSheetRows]
        auditSheetStateID = masterExcelFinal[auditReport_StateID][auditSheetRows]
        if auditSheetFileDate != "":
            try:
                auditSheetFileDate = datetime.datetime.strptime(str(auditSheetFileDate), "%Y-%m-%d %H:%M:%S").strftime("%m-%d-%Y")
            except ValueError:
                logViewer.appendPlainText("Row # [" + str(auditSheetRows + 12) + "] Entity Name[" + str(auditSheetEntityName) + "] StateID[" + str(auditSheetStateID) +"] File Date[" + auditSheetFileDate + "] FileDate has problem!!!")
                QtGui.QGuiApplication.processEvents()
                break
        auditSheetStatus = masterExcelFinal[auditReport_Status][auditSheetRows]
        auditSheetNextARDue = masterExcelFinal[auditReport_NextARDue][auditSheetRows]
        if auditSheetNextARDue != "":
            if auditSheetNextARDue != "***":
                try:
                    auditSheetNextARDue = datetime.datetime.strptime(str(auditSheetNextARDue), "%Y-%m-%d %H:%M:%S").strftime("%m-%d-%Y")
                except ValueError:
                    auditSheetNextARDue = auditSheetNextARDue
            else:
                auditSheetNextARDue = "***"
        else:
            auditSheetNextARDue = ""

        eventRows = 0
        counterEvenSummary = 0
        for eventRows in eventSumReport_FilingEvents.index:
            eventEntityName = eventSumReport_FilingEvents[eventSumReport_EntityName][eventRows]
            eventDomesticJurisdiction = eventSumReport_FilingEvents[eventSumReport_DomesticJurisdiction][eventRows]
            eventJurisdiction = eventSumReport_FilingEvents[eventSumReport_Jurisdiction][eventRows]
            eventDueDate = eventSumReport_FilingEvents[eventSumReport_DueDate][eventRows]
            if eventDueDate != "":
                eventDueDate = datetime.datetime.strptime(str(eventDueDate), "%m-%d-%Y").strftime("%m-%d-%Y")
            eventDomesticJurisdiction = eventDomesticJurisdiction.rstrip()
            eventJurisdiction = eventJurisdiction.rstrip()
            if auditSheetNextARDue == eventDueDate:
                ARDue = 1
            else:
                ARDue = 0

            found_ARDue = 0
            not_found_ARDue = 0
            if auditSheetEntityName == eventEntityName and auditSheetDomesticState == eventDomesticJurisdiction.rstrip() and auditSheetForeignJurisdictions == eventJurisdiction and auditSheetNextARDue != eventDueDate:
                not_found_ARDue = not_found_ARDue + 1
                eventRows = 0
                continue
            if auditSheetEntityName == eventEntityName and auditSheetDomesticState == eventDomesticJurisdiction.rstrip() and auditSheetForeignJurisdictions == eventJurisdiction and auditSheetNextARDue == eventDueDate:
                found_ARDue = found_ARDue + 1
                eventRows = 0
                counterEvenSummary = counterEvenSummary + 1
                break

        if counterEvenSummary == 0:
            if not_found_ARDue >= 1:
                logViewer.appendPlainText("Row # [" + str(auditSheetRows + 12) + "][" + str(auditSheetNextARDue) + "] Next AR Due in Audit Sheet cannot be found in Event Summary Report under Filing Events sheet")
                QtGui.QGuiApplication.processEvents()
                not_found_ARDue = 0
            logViewer.appendPlainText("Row # [" + str(auditSheetRows + 12) + "] Entity Name[" + str(auditSheetEntityName) + "] Domestic State[" + str(auditSheetDomesticState) + "] Jurisdictions[" + str(auditSheetForeignJurisdictions) + "] NextARDue[" + str(auditSheetNextARDue) + "] cannot be found in Event Summary Report under Filing Events sheet.")
            QtGui.QGuiApplication.processEvents()
            counterEvenSummary = 0

    #“Entity Type” should be spelled out
    for masterRows in masterExcelFinal.index:
        masterExcelFinal.fillna('', inplace=True)

        # Entity Type should be spelled out
        masterEntityType = masterExcelFinal.iloc[masterRows, 2]
        masterEntityType = RegEx.search(r'(LLC|LP)', str(masterEntityType))
        if masterEntityType:
            logViewer.appendPlainText("Row# " + str(masterRows + 12) + ": - Found spelled out Entity Type")
            QtGui.QGuiApplication.processEvents()

        # “Domestic State” should be spelled out
        masterDomesticState = masterExcelFinal.iloc[masterRows, 3]
        statesListFile = read_ini_config.readFile(base_path + "\\ForeignStatesList.txt")
        checkValue = RegEx.search(str(masterDomesticState), statesListFile)
        #if not checkValue:
         #   logViewer.appendPlainText("Row# " + str(masterRows + 12) + ": Domestic State is spelled out. Please check")
          #  QtGui.QGuiApplication.processEvents()

        # “Foreign State Registrations” should be spelled out
        masterForeignJur = masterExcelFinal.iloc[masterRows, 4]
        statesListFile = read_ini_config.readFile(base_path + "\\ForeignStatesList.txt")
        checkValue = RegEx.search(str(masterForeignJur), statesListFile)
        #if not checkValue:
         #   logViewer.appendPlainText("Row# " + str(masterRows + 12) + ": Foreign State Registration is spelled out. Please check")
          #  QtGui.QGuiApplication.processEvents()

        # “File Date” column should be in the below format.
        masterFilingDate = masterExcelFinal.iloc[masterRows, 5]
        masterCorrectFilingDate = change_date_format(str(masterFilingDate))
        try:
            CorrectValidDate = is_date(str(masterCorrectFilingDate))
        except:
            continue
        if CorrectValidDate == True:
            logViewer.appendPlainText("Row# " + str(masterRows + 12) + ": Filing date is not in correct format. Should be " + masterCorrectFilingDate)
            QtGui.QGuiApplication.processEvents()

        # “State ID #” - No valid “Date” format e.g 3/14/2019 within the column.
        masterStateID = masterExcelFinal.iloc[masterRows, 6]
        try:
            ValidDate = is_date(str(masterStateID))
        except:
            continue
        if ValidDate == True:
            logViewer.appendPlainText(
                   "Row# " + str(masterRows + 12) + ": Found date format in State ID. Please correct." + masterStateID)
            QtGui.QGuiApplication.processEvents()

        # “Status” only the below statuses are allowed within the column
        masterStatus = masterExcelFinal.iloc[masterRows, 7]
        if masterStatus == "":
            continue
        StatusList = RegEx.search(str(masterStatus), ListStatus)
        #if StatusList is None:
         #   logViewer.appendPlainText("Row# " + str(masterRows + 12) + ": Status is not found in the list of allowed statuses.")
          #  QtGui.QGuiApplication.processEvents()

        # “Status for hCue Build” only the below statuses are allowed within the column
        masterStatushCueBuild = masterExcelFinal.iloc[masterRows, 17]
        if masterStatushCueBuild == "":
            continue
        StatusList = RegEx.search(str(masterStatushCueBuild), ListStatus)
        #if StatusList is None:
            # logViewer.appendPlainText("Row# " + str(masterRows + 12) + ": Status for hCue Build is not found in the list of allowed statuses.")
            # QtGui.QGuiApplication.processEvents()

        # “FYE” below jurisdictions should have no FYE and marked as “Not
        masterFYE = masterExcelFinal.iloc[masterRows, 8]
        # masterDomesticState = masterExcelFinal.iloc[masterRows, 3]
        FYEJurisdictions = read_ini_config.readFile(base_path + "\\JurisdictionsNoFYE.txt")
        jurList = RegEx.search(str(masterForeignJur), FYEJurisdictions)
        if jurList:
            if masterFYE == "FYE":
                logViewer.appendPlainText("Row# " + str(masterRows + 12) + ": Jurisdictions should have no FYE. Should be 'Not Applicable.'")
                QtGui.QGuiApplication.processEvents()

        # Is CT the Registered Agent?


        # Address of Agent
        #masterAgentAddress = masterExcelFinal.iloc[masterRows, 11]
        #if masterAgentAddress == "***":
         #   continue
        #agentAddress = masterAgentAddress.split(',')
        #zipCode = agentAddress[-1]
        #zipCodeValue = RegEx.search(r'[A-Z][A-Z] [0-9][0-9][0-9][0-9][0-9]+', zipCode)
        #if not zipCodeValue:
         #   logViewer.appendPlainText("Row# " + str(masterRows + 1) + ": Zip code is below the minimum 5 digits")
          #  QtGui.QGuiApplication.processEvents()

        # Last AR Due", "Last AR Filed" and "Next AR Due"
        masterEntityType = masterExcelFinal.iloc[masterRows, 2]
        masterDomesticState = masterExcelFinal.iloc[masterRows, 3]
        masterForeignJur = masterExcelFinal.iloc[masterRows, 4]
        masterLastARDue = masterExcelFinal.iloc[masterRows, 12]
        masterLastARFiled = masterExcelFinal.iloc[masterRows, 13]
        masterNextARDue = masterExcelFinal.iloc[masterRows, 14]
        dueDatesPattern = read_ini_config.readLinesOfFile(base_path, "\\Pattern.txt")
        for patternLines in dueDatesPattern:
            EntityType, DomesticState, ForeignJurisdiction, DueStartWith = patternLines.split(';')
            iEntityType, valueEntityType = EntityType.split('=')
            iDomesticState, valueDomesticState = DomesticState.split('=')
            iForeignJurisdiction, valueForeignJurisdiction = ForeignJurisdiction.split('=')
            iDueStartWith, valueDueStartWith = DueStartWith.split('=')
            if valueEntityType == masterEntityType and valueDomesticState == masterDomesticState and valueForeignJurisdiction == masterForeignJur:
                pass
                # masterLastARDue = masterLastARDue[3:]
                # valueDueStartWith = valueDueStartWith[3:]
                # if not masterLastARDue == valueDueStartWith or not masterNextARDue == valueDueStartWith:
                #   logViewer.appendPlainText("Row# " + str(masterRows + 1) + ": Last AR Due or Next AR Due should be :" + valueDueStartWith)
                #  QtGui.QGuiApplication.processEvents()

        # “Would you like assistance changing the agent to CT?”
        masterAssistanceToCT = masterExcelFinal.iloc[masterRows, 21]
        masterCTAgent = masterExcelFinal.iloc[masterRows, 9]
        if masterCTAgent == "No":
            if masterAssistanceToCT != "":
                logViewer.appendPlainText("Row# " + str(
                    masterRows + 12) + ": Value of “Is CT the Registered Agent?” is No, the entry 'Would you like assistance changing the agent to CT?' column should be blank")
                QtGui.QGuiApplication.processEvents()
        if masterCTAgent == "Yes":
            if masterAssistanceToCT != "Not Applicable":
                logViewer.appendPlainText("Row# " + str(
                    masterRows + 12) + ": Value of “Is CT the Registered Agent?” is Yes, the entry 'Would you like assistance changing the agent to CT?' column should be 'Not Applicable'")
                QtGui.QGuiApplication.processEvents()

        # “Would you like me to file the past due annual report(s)?”
        masterPastDueAnnual = masterExcelFinal.iloc[masterRows, 22]
        # need validation

        # “Would you like assistance reinstating?”
        masterPastDueAnnual = masterExcelFinal.iloc[masterRows, 23]
        # need validation

    logViewer.appendPlainText("\r\nDone. Please see above logs in " + logPath + " CTCOB17-AuditTemplateValidation.logs")
    QtGui.QGuiApplication.processEvents()
    showDialog("Done", "Done Processing")
    read_ini_config.writeFileWrite(logPath + "CTCOB17-AuditTemplateValidation.logs", logViewer.toPlainText())


def change_date_format(dt):
    return RegEx.sub(r'(\d{4})-(\d{1,2})-(\d{1,2})', '\\2/\\3/\\1', dt)


def manualBrowsePathMaster():
    options = QFileDialog.Options()
    options |= QFileDialog.DontUseNativeDialog
    fileName, _ = QFileDialog.getOpenFileName(None, "Browse Excel File", inputPath,
                                              "Excel File Macro(*.xlsm);;Excel File(*.xlsx);;Excel File (*.xls)", options=options)
    if fileName:
        fileName = os.path.normpath(fileName)
        masterPathText.setText(fileName)


def manualBrowsePathCorporate():
    options = QFileDialog.Options()
    options |= QFileDialog.DontUseNativeDialog
    fileName, _ = QFileDialog.getOpenFileName(None, "Browse Corporate Summary File", inputPath,
                                              "Excel File Macro(*.xlsm);;Excel File(*.xlsx);;Excel File (*.xls)", options=options)
    if fileName:
        fileName = os.path.normpath(fileName)
        corporatePathText.setText(fileName)

def manualBrowsePathEvent():
    options = QFileDialog.Options()
    options |= QFileDialog.DontUseNativeDialog
    fileName, _ = QFileDialog.getOpenFileName(None, "Browse Event Summary File", inputPath,
                                              "Excel File Macro(*.xlsm);;Excel File(*.xlsx);;Excel File (*.xls)", options=options)
    if fileName:
        fileName = os.path.normpath(fileName)
        eventPathText.setText(fileName)


def showDialog(textMessage, messageTitle):
    msg = QMessageBox()
    msg.setIcon(QMessageBox.Information)
    msg.setText(textMessage)
    msg.setWindowTitle(messageTitle)
    msg.setModal(True)
    msg.setStandardButtons(QMessageBox.Ok)
    msg.setDefaultButton(QMessageBox.Ok)
    msg.exec_()
    return


# connect butttons
validateButton.clicked.connect(startValidations)
browseMasterPath.clicked.connect(manualBrowsePathMaster)
browseCorporatePath.clicked.connect(manualBrowsePathCorporate)
browseEventPath.clicked.connect(manualBrowsePathEvent)

# Show the window and run the app
win.show()
app.exec_()
