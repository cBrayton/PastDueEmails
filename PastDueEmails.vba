Option Compare Database
Private records As Recordset
Private overPastDue As New Collection
Private PDList As String
Private SupervisorList As String
Private outlookApp As Object
Private outlookNamespace As Object
Private dayThreshold As Integer


Sub Filter90Days() 'Filters a past due report to delete anything less than 90 days past due, so only supervisors with employees over 90 days past due get an email.
'***Note***This deletes records less than 90 days past due, so the report will need to be added to access again if those records are needed.
'This function is now improved to be modifiable based on the form. The 90 day threshold can now be changed.
Dim today As Date
Dim daysPastDue As Integer
Dim overPastDueEmp As String
'Dim records As Recordset
Dim PastDueReport As String
'Dim dayThreshold As Integer
PDList = List9.Value
dayThreshold = Text16.Value
PastDueReport = CStr(PDList)
'PastDueReport = InputBox("What is the name of the past due report in Access?", , "ISOExport_PastDue")
DoCmd.RunSQL ("SELECT " & SupervisorList & "_1.EMP_ID, " & SupervisorList & "_1.EMPLOYEE_NAME, " & SupervisorList & ".EMP_ID, " & SupervisorList & ".EMPLOYEE_NAME INTO SupervisorTable FROM " & SupervisorList & " INNER JOIN " & SupervisorList & " AS " & SupervisorList & "_1 ON " & SupervisorList & ".SUP_NAME = " & SupervisorList & "_1.EMPLOYEE_NAME ORDER BY " & SupervisorList & "_1.EMP_ID;")
DoCmd.RunSQL ("select SupervisorTable." & SupervisorList & "_1_EMP_ID, " & PastDueReport & ".EMP_NAME, " & PastDueReport & ".AGING_DAYS into PastDueTable from " & PastDueReport & " left join SupervisorTable on " & PastDueReport & ".EMP_ID = SupervisorTable." & SupervisorList & "_EMP_ID WHERE ((" & PastDueReport & ".[STATUS]='MISSING')) OR ((" & PastDueReport & ".[STATUS]='RETRAIN')) OR ((" & PastDueReport & ".[STATUS]='REVISION')) order by SupervisorTable." & SupervisorList & "_1_EMP_ID, " & PastDueReport & ".EMP_NAME")
Set records = CurrentDb.OpenRecordset("select * from PastDueTable")
'Set records = CurrentDb.OpenRecordset("select SupervisorTable." & SupervisorList & "_1_EMP_ID, " & PastDueReport & ".EMP_NAME, " & PastDueReport & ".AGING_DAYS into PastDueTable from " & PastDueReport & " left join SupervisorTable on " & PastDueReport & ".EMP_ID = SupervisorTable." & SupervisorList & "_EMP_ID order by SupervisorTable." & SupervisorList & "_1_EMP_ID")
today = Date

Do While Not records.EOF
   
    daysPastDue = records(2)
    If daysPastDue <= dayThreshold Then 'Removes any records with 90 or fewer days past due. Change the 90 here to search for a different number of days
        records.Delete
    ElseIf dayThreshold < 0 And daysPastDue > 0 Then
        records.Delete
    End If
    If daysPastDue > dayThreshold * 2 And dayThreshold > 0 Then 'Adds any records with 181 or more days past due to the overPastDue collection. Change the 180 here to email one over the supervisor at a lower number of days past due.
        overPastDueEmp = records(1)
        overPastDue.Add (overPastDueEmp)
    End If
    If IsNull(records(0)) Then
        records.Edit
        records(0) = ""
        records.Update
    End If
    records.MoveNext
   
Loop
records.MoveFirst
'Filter90Days = records
End Sub

Function SupervisorLookup(EmpID As String) As String 'Takes an employee's ID, then looks up and returns their supervisor's ID.
SupervisorList = List11.Value
SupervisorList = CStr(SupervisorList)
'SupervisorList = InputBox("What is the name of the Supervisor list in Access?", , "ISOExport_Sups")
Dim SupEmpID As Recordset
Dim SupID As String
Dim SupName As String
Set SupEmpID = CurrentDb.OpenRecordset("select EMP_ID, EMPLOYEE_NAME, SUP_NAME from " & SupervisorList & " order by EMP_ID")
Do While Not SupEmpID.EOF 'Loops through all records to find the employee ID in the employee supervisor database
   
    If EmpID = SupEmpID(0) And SupEmpID(2) <> "" Then
        SupName = SupEmpID(2)
        SupEmpID.MoveFirst
       
        Do While Not SupEmpID.EOF 'Loops through all records to match the supervisor's name to his/her ID.
       
            If SupName = SupEmpID(1) Then
                SupervisorLookup = SupEmpID(0)
                Exit Function
            Else
                SupEmpID.MoveNext
                If SupEmpID.EOF Then     'Sends an email to cBrayton (author) informing of an exception in the find supervisor function
'''''                    DoCmd.SendObject acSendNoObject, Null, acFormatHTML, "author@domain.com", "", "", "One-over supervisor isn't in the employee supervisor database", "The employee with ID code" & EmpID & "has a past due employee and their one-over supervisor isn't in the database", False, ""
                    Exit Function
                End If
            End If
        Loop
    Else
        SupEmpID.MoveNext
        If SupEmpID.EOF Then     'Sends an email to cBrayton (author) informing of an exception where the employee with past due training is not in the supervisor lookup database; a new supervisor report will need to be run and stored in the database.
'''''            DoCmd.SendObject acSendNoObject, Null, acFormatHTML, "author@domain.com", "supervisors_email_address@domain.com", "", "Supervisor with a Past Due employee is not in employee-supervisor DB or doesn't have a supervisor in the database.", "The supervisor with ID code " & EmpID & " has a past due employee and is not in the employee-supervisor database", False, ""
            Exit Function
        End If
    End If
Loop

End Function

Sub MassEmail()
' Emails supervisors informing them of their employees with training over 90 days past due, can optionally also email the supervisor's supervisor about the past due training
' The default functionality is to email the supervisor's supervisor if an employee has training over 180 days past due
' The system is setup to send either HTML or plain text emails. Comment out the lines with DoCmd to send HTML emails, or comment out the lines with SendEmail to send plain text emails
Dim DoNOTSendTable As DAO.Recordset
'Dim DoNOTSendList As Variant(4)(DoNOTSentTable.RecordCount) 'This is a two dimensional array with people whose supervisors should not be emailed with the ID number to use as an email address paired with the name.
'People on the DoNOTSendList will receive an email informing them they have past due training.
'DoNOTSendList = Array(Array("", ""))
Set DoNOTSendTable = CurrentDb.OpenRecordset("select * from DoNotSend")
Dim DoNOTSendSize As Integer
DoNOTSendSize = DoNOTSendTable.RecordCount()
DoNOTSendList = DoNOTSendTable.GetRows(8) '(8) changed from (DoNOTSendSize)
Dim emailBody As String
Dim emailFooter As String
'**************************************************************************************************
'Edit the following strings to change the text before and after the list of employees in the email.
'**************************************************************************************************
emailBody = "One or more of your employees have training >90 days Past Due. " _
    & "Please take this opportunity to follow-up with them to ensure they complete their assigned training. " _
    & "Per SOPs, it is your responsibility to ensure your staff remains current on assigned training. " _
    & vbCrLf & vbCrLf & "The following employee(s) have Past Due (PD) training (* indicates one or more courses >180 days PD):" _
    & vbCrLf & vbCrLf & "Employee Name:              # of Courses 90+ Days Past Due:" & vbCrLf
emailFooter = vbCrLf & "Visit ISOtrain for more information." & vbCrLf
'**************************************************************************************************
'**************************************************************************************************
Dim HTMLemailBody As String
Dim HTMLemailFooter As String
Dim HTMLemployees As String
'**************************************************************************************************
'Edit the following strings to change the HTML text before and after the list of employees in the email.
'Comment out the lines below if sending a plain text email.
'**************************************************************************************************
If dayThreshold < 0 Then
    HTMLemailBody = "<p>One or more of your employees has training due within the next " & Right(CStr(dayThreshold), Len(CStr(dayThreshold)) - 1) & " days (in addition to any past due training s/he may have).</p>" _
        & "<p>Please take this opportunity to access ISOtrain and help ensure training is completed on-time. Per LA26000119, <i>Training and Qualification Requirements</i> (§6.1.1.6), it is your responsibility to ensure your staff remains current on assigned training.</p>" _
        & "<p>In most cases, it is not appropriate to remove a module because training is about to become past due.</p>" _
        & "<p>The following employee(s) have Upcoming Training:</p>"
    HTMLemailFooter = "Visit ISOtrain for more information."
Else
    HTMLemailBody = "<p>One or more of your employees have training >" & dayThreshold & " days Past Due. " _
        & "Please take this opportunity to follow-up with them to ensure they complete their assigned training. " _
        & "Per SOPs, it is your responsibility to ensure your staff remains current on assigned training.</p> " _
        & "<p>The following employee(s) have Past Due (PD) training (* indicates one or more courses >" & dayThreshold * 2 & " days PD):</p>"
        '& vbCrLf & vbCrLf & "Employee Name:              # of Courses 90+ Days Past Due:" & vbCrLf
    HTMLemailFooter = "Visit ISOtrain for more information."
End If
'**************************************************************************************************
'**************************************************************************************************
Dim HTMLemailBodyDefault As String
HTMLemailBodyDefault = HTMLemailBody
Dim emailAddress As String
Dim employees As String
Dim lastemployee As String
Dim emailText As String
Dim counter As Integer
Dim counterText As String
Dim spaces As String
Dim supCCaddress As String
Dim mostPastDues As String
Dim maxPastDue As Integer
Dim DebugCC As String
Dim cumulativeEmployeeList As String
Dim HTMLcumulativeEmployeeList As String
Dim employeeCounter As Integer
Dim supCheck As Boolean
'Dim records As Recordset
'Set records = CurrentDb.OpenRecordset("select Sup_Email, Emp_Name from EmpSupCells order by Sup_Email")
Filter90Days
maxPastDue = 0
Do While Not records.EOF
    employees = "" 'Comment this line to send a cumulative list of employees (for debugging only; should also comment the send lines in the loop)
    HTMLemployees = ""
    counter = 0
    supCheck = False
    lastemployee = records(1)
    Do
        Do
      
            emailAddress = records(0)
            counter = counter + 1
            lastemployee = records(1)
            records.MoveNext
            If records.EOF Then
                Exit Do
            End If
        Loop While lastemployee = records(1)
       
        'employees = employees & lastemployee & "has " & counter & " course(s) over 90 days past due; " & vbCrLf 'Uncomment this line for a different email text format.
        employeeCounter = employeeCounter + 1
        counterText = counter
        If counter > maxPastDue Then
            maxPastDue = counter
            mostPastDues = lastemployee
        End If
        For Each emp In overPastDue 'Adds * next to the number of courses past due to indicate that at least one course is over 180 days past due
            If InStr(1, lastemployee, emp) And Not CBool(InStr(1, counterText, "*")) Then
                counterText = counterText & "*" 'Can change this to reorder the * and number or change the asterisks (if changed code needs some changes too).
            End If
        Next
        For i = 1 To 7 - Int((Len(lastemployee)) / 7) 'Can edit this to modifiy the spacing between employee name and course count. This is for a 12 point New Courier font for plaintext emails.
            spaces = vbTab & spaces
        Next i
        spaces = "  " & spaces
        employees = employees & lastemployee & spaces & counterText & vbCrLf
        HTMLemployees = HTMLemployees & "<tr style='mso-yfti-irow:" & employeeCounter & "'><td width=319 valign=top style='width:239.4pt;padding:0in 5.4pt 0in 5.4pt'><p>" & lastemployee & "<o:p></o:p></p></td><td width=319 valign=top style='width:239.4pt;padding:0in 5.4pt 0in 5.4pt'><p align=right style='text-align:right'>" & counterText
        spaces = ""
        counter = 0
        'If records.EOF Then 'Moved below the following if block
        '    Exit Do
        'End If
        If InStr(1, counterText, "*") And Not supCheck Then 'Looks for a * in the employees' past due text, and sends the email to the supervisor's supervisor if an employee has training over 180 days past due.
            supCCaddress = SupervisorLookup(emailAddress) 'Comment this line to send an email only to the employees' direct supervisor; uncomment this line to prevent sending an email to supervisor and supervisor's supervisor.
            HTMLemployees = HTMLemployees & "<o:p></o:p></p></td></tr>"
            supCheck = True
        ElseIf InStr(1, counterText, "*") Then
            HTMLemployees = HTMLemployees & "<o:p></o:p></p></td></tr>"
        Else
            HTMLemployees = HTMLemployees & "<span style='mso-ascii-font-family:Times New Roman;mso-hansi-font-family:Times New Roman;color:white;mso-themecolor:background1'>*<o:p></o:p></span><o:p></o:p></p></td></tr>"
        End If
        If records.EOF Then
            Exit Do
        End If
       
    Loop While records(0) = emailAddress
    'emailaddress = "author"     'Uncomment this line for debugging to send all emails to one address (ie debugger/author)
    'supCCaddress = "supervisors_email_address"  'Uncomment this line for debugging to cc all emails to one address (ie debugger/author's supervisor)
    emailText = emailBody & employees & emailFooter
    HTMLemailBody = HTMLemailBodyDefault
    If emailAddress = "" Then 'Emails the author and supervisor if the employees don't have direct supervisors in ISOtrain.
        emailAddress = "author"
        supCCaddress = "supervisors_email_address"
        emailText = "These Employees don't have a direct supervisor." & vbCrLf & vbCrLf & emailText
        HTMLemailBody = "<p>These Employees don't have a direct supervisor.</p>" & HTMLemailBody
    End If
    For i = 0 To DoNOTSendTable.RecordCount - 1 'Can make this dynamic to run through the whole array in case size changes
        If CBool(InStr(1, emailText, DoNOTSendList(2, i))) Or emailAddress = DoNOTSendList(1, i) Or emailAddress = DoNOTSendList(2, i) Or emailAddress = DoNOTSendList(3, i) Or emailAddress = DoNOTSendList(4, i) Then
            emailAddress = DoNOTSendList(1, i)
            supCCaddress = ""
        End If
    Next i
    'emailText = emailaddress & supCCaddress & emailText 'Uncomment this line for debugging to see the original recipient
    'HTMLemailBody = "<p>Original recipient: " & emailAddress & " and original CC: " & supCCaddress & "</p>" & HTMLemailBody 'Uncomment this line for debugging to see the original recipient for HTMl emails.
    'emailAddress = "author"     'Uncomment this line for debugging to send all emails to one address (ie debugger/author) currently cBrayton
    'supCCaddress = "supervisors_email_address"           'Uncomment this line for debugging to ensure no one is CC'd on the test emails
    'supCCaddress = ""
    If supCCaddress <> "" Then 'The SendEmail code should send HTML formatted emails for better spacing of the names.
        supCCaddress = supCCaddress & "@domain.com" & "; author@domain.com; supervisors_email_address" 'Uncomment this line for debugging to cc the author and supervisor on all emails w/ cc already
        SendEmail HTMLemailBody, HTMLemployees, HTMLemailFooter, emailAddress, supCCaddress
        'DoCmd.SendObject acSendNoObject, Null, acFormatHTML, emailAddress & "@domain.com", supCCaddress & "@domain.com", "", "NOTIFICATION: Aging Past Due Training", emailText, False, ""
    Else
        DebugCC = "author@domain.com; supervisors_email_address" 'Uncomment this line for debugging to cc the author and supervisor on all emails w/o cc already
        SendEmail HTMLemailBody, HTMLemployees, HTMLemailFooter, emailAddress, DebugCC
        'DoCmd.SendObject acSendNoObject, Null, acFormatHTML, emailAddress & "@domain.com", DebugCC &"@domain.com", "", "NOTIFICATION: Aging Past Due Training", emailText, False, ""
    End If
    supCCaddress = ""
    employeeCounter = 0
    cumulativeEmployeeList = cumulativeEmployeeList & employees
    HTMLcumulativeEmployeeList = HTMLcumulativeEmployeeList & HTMLemployees
Loop
emailText = emailBody & cumulativeEmployeeList & emailFooter
SendEmail HTMLemailBody, HTMLcumulativeEmployeeList, HTMLemailFooter, "author", "supervisors_email_address"
'DoCmd.SendObject acSendNoObject, Null, acFormatHTML, "author@domain.com", "supervisors_email_address@domain.com", "", "Compiled list of employees with Past Due training", emailText, False, ""
'Sends an email to the author reporting the person with the most courses over 90 days past due and the number of courses
emailText = mostPastDues & " has the most courses over 90 days past due at " & CStr(maxPastDue)
DoCmd.SendObject acSendNoObject, Null, acFormatTXT, "author@domain.com", "", "", "Employee with the most training over 90 days past due.", emailText, False, ""
DoCmd.Close acTable, "PastDueTable", acSaveNo
Set records = Nothing
CurrentDb.TableDefs.Delete "PastDueTable"
CurrentDb.TableDefs.Delete "SupervisorTable"
End Sub


Public Sub Command0_Click()
    PDList = List9.Value
    SupervisorList = List11.Value
    dayThreshold = Text16.Value
    MassEmail
End Sub


Private Sub Command13_Click()
    '*******Imports the Past Due report and names it ISOExport_PDReportMM/DD/YYYY
    Dim strXML As String
    Dim i, j, k As Integer
    Dim strReplace As String
    Dim ImportFile As Variant
    Dim strImportFile As String
    Dim destinationXML As String
    Dim rangeXML As String
    Do 'Loops until a file with ISOExport_P is in the file path
        Set ImportFile = Application.FileDialog(3)
        ImportFile.show
        strImportFile = ImportFile.SelectedItems(1)
    Loop While Not CBool(InStr(1, strImportFile, "ISOExport_P"))
    strXML = Access.CurrentProject.ImportExportSpecifications(3).XML
    destinationXML = "ISOExport_PDReport" & Trim(Left(CStr(FileDateTime(strImportFile)), InStr(CStr(FileDateTime(strImportFile)), " ")))
    destinationXML = Replace(destinationXML, "/", "_")
    'rangeXML = "ISOExport_PDReport" & Left(CStr(FileDateTime(strImportFile)), 10) & "$"
    rangeXML = ""
    '****Code Below taken from http://www.utteraccess.com/forum/VBA-Manipulation-Saved-E-t1990584.html
    'Find the delimiters of the current path string
    i = InStr(strXML, "ImportExportSpecification Path =")
    'First double quote after that
    i = InStr(i, strXML, Chr(34))
    'Second double quote after that
    j = InStr(i + 1, strXML, Chr(34))
    strReplace = Mid(strXML, i + 1, j - i - 1)
    strXML = Replace(strXML, strReplace, strImportFile)
    'Find the delimiters of the current path string for the Destination
    i = InStr(strXML, "Destination=")
    'First double quote after that
    i = InStr(i, strXML, Chr(34))
    'Second double quote after that
    j = InStr(i + 1, strXML, Chr(34))
    strReplace = Mid(strXML, i + 1, j - i - 1)
    strXML = Replace(strXML, strReplace, destinationXML)
    ''Find the delimiters of the current range string then removes the range string
    'i = InStr(strXML, "Range=")
    ''First double quote after that
    'k = InStr(i, strXML, Chr(34))
    ''Second double quote after that
    'j = InStr(k + 1, strXML, Chr(34))
    'strReplace = Mid(strXML, i, j - i + 1)
    'strXML = Replace(strXML, strReplace, rangeXML)
    Access.CurrentProject.ImportExportSpecifications(3).XML = strXML
    Access.CurrentProject.ImportExportSpecifications(3).Execute
    List9.Requery

End Sub

Private Sub Command14_Click()

    '****Imports the Supervisor List report and names it ISOExport_SupsMM/DD/YYYY
    Dim strXML As String
    Dim i, j As Integer
    Dim strReplace As String
    Dim ImportFile As Variant
    Dim strImportFile As String
    Dim destinationXML As String
    Dim rangeXML As String
    Do 'Loops until a file with ISOExport_S is in the file path
    Set ImportFile = Application.FileDialog(3)
    ImportFile.show
    strImportFile = ImportFile.SelectedItems(1)
    Loop While Not CBool(InStr(1, strImportFile, "ISOExport_S"))
    strXML = Access.CurrentProject.ImportExportSpecifications(1).XML
    destinationXML = "ISOExport_Sups" & Trim(Left(CStr(FileDateTime(strImportFile)), InStr(CStr(FileDateTime(strImportFile)), " ")))
    destinationXML = Replace(destinationXML, "/", "_")
    rangeXML = "ISOExport_Sups" & Left(CStr(FileDateTime(strImportFile)), 10) & "$"
    '****Code Below taken from http://www.utteraccess.com/forum/VBA-Manipulation-Saved-E-t1990584.html
    'Find the delimiters of the current path string
    i = InStr(strXML, "ImportExportSpecification Path =")
    'First double quote after that
    i = InStr(i, strXML, Chr(34))
    'Second double quote after that
    j = InStr(i + 1, strXML, Chr(34))
    strReplace = Mid(strXML, i + 1, j - i - 1)
    strXML = Replace(strXML, strReplace, strImportFile)
    'Find the delimiters of the current path string for the Destination
    i = InStr(strXML, "Destination=")
    'First double quote after that
    i = InStr(i, strXML, Chr(34))
    'Second double quote after that
    j = InStr(i + 1, strXML, Chr(34))
    strReplace = Mid(strXML, i + 1, j - i - 1)
    strXML = Replace(strXML, strReplace, destinationXML)
    ''Range entry isn't used in the supervisor import
    ''Find the delimiters of the current path string
    'i = InStr(strXML, "Range=")
    ''First double quote after that
    'i = InStr(i, strXML, Chr(34))
    ''Second double quote after that
    'j = InStr(i + 1, strXML, Chr(34))
    'strReplace = Mid(strXML, i + 1, j - i - 1)
    'strXML = Replace(strXML, strReplace, rangeXML)
    Access.CurrentProject.ImportExportSpecifications(1).XML = strXML
    Access.CurrentProject.ImportExportSpecifications(1).Execute
    List11.Requery
   
End Sub

Private Sub Command15_Click()

Dim top12 As String
Dim stringSQL As String
PDList = List9.Value
PDList = CStr(PDList)
top12 = "Top12_PD" & Right(PDList, 10)
SupervisorList = List11.Value
SupervisorList = CStr(SupervisorList)
stringSQL = "SELECT TOP 12 Round(Now()-[DUE_DATE],0) AS DaysOld, " & SupervisorList & ".EMPLOYEE_NAME AS Supervisor, " & PDList & ".EMP_ID, " & PDList & ".COURSE_CODE, " & PDList & ".DESCRIPTION " _
& "INTO " & top12 & " " _
& "FROM " & PDList & " LEFT JOIN " & SupervisorList & " ON " & PDList & ".SUPERVISOR_CODE = " & SupervisorList & ".EMP_ID " _
& "ORDER BY " & PDList & ".DUE_DATE;"
DoCmd.SetWarnings False
DoCmd.RunSQL (stringSQL)
DoCmd.SetWarnings True

End Sub

Private Sub List9_GotFocus()
    List9.Requery
    List11.Requery
End Sub

Private Sub List9_LostFocus()
    List9.Requery
    List11.Requery
End Sub

Private Sub List11_GotFocus()
    List9.Requery
    List11.Requery
End Sub

Private Sub List11_LostFocus()
    List9.Requery
    List11.Requery
End Sub

Sub InitOutlook()
    ' Initialize a session in Outlook
    Set outlookApp = CreateObject("Outlook.Application")
   
    'Return a reference to the MAPI layer
    Set outlookNamespace = outlookApp.GetNamespace("MAPI")
   
    'Let the user logon to Outlook with the
    'Outlook Profile dialog box
    'and then create a new session
    outlookNamespace.Logon , , True, False
End Sub
 
Sub Cleanup()
    ' Clean up public object references.
    Set outlookNamespace = Nothing
    Set outlookApp = Nothing
End Sub
 
Sub SendEmail(emailBody, employeeList, emailFooter, emailAddress, Optional ccAddress)
    'Takes email addresses as employee ID numbers or First_Last names then adds @domain to the ID/Name
    Dim mailItem As Object
    If dayThreshold < 0 Then
        second_column_header = "# of Courses Due Within " & Right(CStr(dayThreshold), Len(CStr(dayThreshold)) - 1) & " Days:"
    Else
        second_column_header = "# of Courses " & dayThreshold & "+ Days Past Due:"
    End If
    InitOutlook
    Set mailItem = outlookApp.createitem(olMailItem)
    mailItem.To = emailAddress & "@domain.com"
    If ccAddress <> "" Then
        ccAddress = ccAddress & "@domain.com"
    End If
    mailItem.CC = ccAddress
    If dayThreshold < 0 Then
    mailItem.Subject = "NOTIFICATION: Your employees with training due within " & Right(CStr(dayThreshold), Len(CStr(dayThreshold)) - 1) & " days."
    Else
    mailItem.Subject = "NOTIFICATION: Aging Past Due Training"
    End If
    mailItem.BodyFormat = 2
    'The email body below was taken straight from outlook. Create a message draft with the desired formatting, then right click and view source. Copy the source over to a word doc, then find/replace all " to """" and add line breaks at the end of each line, and concatenate the body text to itself every 24 lines.
    'Minor changes can be done on this template without needing to do the above steps ^.
    mailItem.HTMLBody = "<html xmlns:v=""""urn:schemas-microsoft-com:vml"""" xmlns:o=""""urn:schemas-microsoft-com:office:office"""" xmlns:w=""""urn:schemas-microsoft-com:office:word"""" xmlns:m=""""http://schemas.microsoft.com/office/2004/12/omml"""" xmlns=""""http://www.w3.org/TR/REC-html40""""><head><meta http-equiv=Content-Type content=""""text/html; charset=us-ascii""""><meta name=ProgId content=Word.Document><meta name=Generator content=""""Microsoft Word 14""""><meta name=Originator content=""""Microsoft Word 14""""><link rel=File-List href=""""cid:filelist.xml@01CF02E8.0A7E95D0""""><!--[if gte mso 9]><xml>" _
        & "<o:OfficeDocumentSettings>" _
        & "<o:AllowPNG/>" _
        & "</o:OfficeDocumentSettings>" _
        & "</xml><![endif]--><!--[if gte mso 9]><xml>" _
        & "<w:WordDocument>" _
        & "<w:SpellingState>Clean</w:SpellingState>" _
        & "<w:TrackMoves/>" _
        & "<w:TrackFormatting/>" _
        & "<w:EnvelopeVis/>" _
        & "<w:PunctuationKerning/>" _
        & "<w:ValidateAgainstSchemas/>" _
        & "<w:SaveIfXMLInvalid>false</w:SaveIfXMLInvalid>" _
        & "<w:IgnoreMixedContent>false</w:IgnoreMixedContent>" _
        & "<w:AlwaysShowPlaceholderText>false</w:AlwaysShowPlaceholderText>" _
        & "<w:DoNotPromoteQF/>" _
        & "<w:LidThemeOther>EN-US</w:LidThemeOther>" _
        & "<w:LidThemeAsian>X-NONE</w:LidThemeAsian>" _
        & "<w:LidThemeComplexScript>X-NONE</w:LidThemeComplexScript>" _
        & "<w:Compatibility>" _
        & "<w:BreakWrappedTables/>" _
        & "<w:SnapToGridInCell/>" _
        & "<w:WrapTextWithPunct/>" _
        & "<w:UseAsianBreakRules/>" _
        & "<w:DontGrowAutofit/>"
    mailItem.HTMLBody = mailItem.HTMLBody _
        & "<w:SplitPgBreakAndParaMark/>" _
        & "<w:EnableOpenTypeKerning/>" _
        & "<w:DontFlipMirrorIndents/>" _
        & "<w:OverrideTableStyleHps/>" _
        & "</w:Compatibility>" _
        & "<m:mathPr>" _
        & "<m:mathFont m:val=""""Cambria Math""""/>" _
        & "<m:brkBin m:val=""""before""""/>" _
        & "<m:brkBinSub m:val=""""&#45;-""""/>" _
        & "<m:smallFrac m:val=""""off""""/>" _
        & "<m:dispDef/>" _
        & "<m:lMargin m:val=""""0""""/>" _
        & "<m:rMargin m:val=""""0""""/>" _
        & "<m:defJc m:val=""""centerGroup""""/>" _
        & "<m:wrapIndent m:val=""""1440""""/>" _
        & "<m:intLim m:val=""""subSup""""/>" _
        & "<m:naryLim m:val=""""undOvr""""/>" _
        & "</m:mathPr></w:WordDocument>" _
        & "</xml><![endif]--><!--[if gte mso 9]><xml>" _
        & "<w:LatentStyles DefLockedState=""""false"""" DefUnhideWhenUsed=""""true"""" DefSemiHidden=""""true"""" DefQFormat=""""false"""" DefPriority=""""99"""" LatentStyleCount=""""267"""">" _
        & "<w:LsdException Locked=""""false"""" Priority=""""0"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" QFormat=""""true"""" Name=""""Normal""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""9"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" QFormat=""""true"""" Name=""""heading 1""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""9"""" QFormat=""""true"""" Name=""""heading 2""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""9"""" QFormat=""""true"""" Name=""""heading 3""""/>"
    mailItem.HTMLBody = mailItem.HTMLBody _
        & "<w:LsdException Locked=""""false"""" Priority=""""9"""" QFormat=""""true"""" Name=""""heading 4""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""9"""" QFormat=""""true"""" Name=""""heading 5""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""9"""" QFormat=""""true"""" Name=""""heading 6""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""9"""" QFormat=""""true"""" Name=""""heading 7""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""9"""" QFormat=""""true"""" Name=""""heading 8""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""9"""" QFormat=""""true"""" Name=""""heading 9""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""39"""" Name=""""toc 1""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""39"""" Name=""""toc 2""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""39"""" Name=""""toc 3""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""39"""" Name=""""toc 4""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""39"""" Name=""""toc 5""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""39"""" Name=""""toc 6""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""39"""" Name=""""toc 7""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""39"""" Name=""""toc 8""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""39"""" Name=""""toc 9""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""35"""" QFormat=""""true"""" Name=""""caption""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""10"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" QFormat=""""true"""" Name=""""Title""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""1"""" Name=""""Default Paragraph Font""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""11"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" QFormat=""""true"""" Name=""""Subtitle""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""22"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" QFormat=""""true"""" Name=""""Strong""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""20"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" QFormat=""""true"""" Name=""""Emphasis""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""59"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Table Grid""""/>" _
        & "<w:LsdException Locked=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Placeholder Text""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""1"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" QFormat=""""true"""" Name=""""No Spacing""""/>"
    mailItem.HTMLBody = mailItem.HTMLBody _
        & "<w:LsdException Locked=""""false"""" Priority=""""60"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Light Shading""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""61"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Light List""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""62"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Light Grid""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""63"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Medium Shading 1""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""64"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Medium Shading 2""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""65"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Medium List 1""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""66"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Medium List 2""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""67"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Medium Grid 1""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""68"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Medium Grid 2""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""69"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Medium Grid 3""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""70"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Dark List""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""71"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Colorful Shading""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""72"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Colorful List""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""73"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Colorful Grid""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""60"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Light Shading Accent 1""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""61"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Light List Accent 1""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""62"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Light Grid Accent 1""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""63"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Medium Shading 1 Accent 1""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""64"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Medium Shading 2 Accent 1""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""65"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Medium List 1 Accent 1""""/>" _
        & "<w:LsdException Locked=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Revision""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""34"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" QFormat=""""true"""" Name=""""List Paragraph""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""29"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" QFormat=""""true"""" Name=""""Quote""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""30"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" QFormat=""""true"""" Name=""""Intense Quote""""/> "
    mailItem.HTMLBody = mailItem.HTMLBody _
        & "<w:LsdException Locked=""""false"""" Priority=""""66"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Medium List 2 Accent 1""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""67"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Medium Grid 1 Accent 1""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""68"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Medium Grid 2 Accent 1""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""69"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Medium Grid 3 Accent 1""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""70"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Dark List Accent 1""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""71"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Colorful Shading Accent 1""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""72"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Colorful List Accent 1""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""73"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Colorful Grid Accent 1""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""60"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Light Shading Accent 2""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""61"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Light List Accent 2""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""62"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Light Grid Accent 2""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""63"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Medium Shading 1 Accent 2""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""64"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Medium Shading 2 Accent 2""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""65"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Medium List 1 Accent 2""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""66"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Medium List 2 Accent 2""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""67"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Medium Grid 1 Accent 2""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""68"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Medium Grid 2 Accent 2""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""69"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Medium Grid 3 Accent 2""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""70"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Dark List Accent 2""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""71"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Colorful Shading Accent 2""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""72"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Colorful List Accent 2""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""73"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Colorful Grid Accent 2""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""60"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Light Shading Accent 3""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""61"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Light List Accent 3""""/> "
    mailItem.HTMLBody = mailItem.HTMLBody _
        & "<w:LsdException Locked=""""false"""" Priority=""""62"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Light Grid Accent 3""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""63"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Medium Shading 1 Accent 3""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""64"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Medium Shading 2 Accent 3""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""65"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Medium List 1 Accent 3""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""66"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Medium List 2 Accent 3""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""67"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Medium Grid 1 Accent 3""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""68"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Medium Grid 2 Accent 3""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""69"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Medium Grid 3 Accent 3""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""70"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Dark List Accent 3""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""71"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Colorful Shading Accent 3""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""72"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Colorful List Accent 3""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""73"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Colorful Grid Accent 3""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""60"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Light Shading Accent 4""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""61"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Light List Accent 4""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""62"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Light Grid Accent 4""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""63"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Medium Shading 1 Accent 4""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""64"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Medium Shading 2 Accent 4""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""65"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Medium List 1 Accent 4""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""66"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Medium List 2 Accent 4""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""67"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Medium Grid 1 Accent 4""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""68"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Medium Grid 2 Accent 4""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""69"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Medium Grid 3 Accent 4""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""70"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Dark List Accent 4""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""71"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Colorful Shading Accent 4""""/> "
    mailItem.HTMLBody = mailItem.HTMLBody _
        & "<w:LsdException Locked=""""false"""" Priority=""""72"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Colorful List Accent 4""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""73"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Colorful Grid Accent 4""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""60"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Light Shading Accent 5""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""61"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Light List Accent 5""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""62"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Light Grid Accent 5""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""63"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Medium Shading 1 Accent 5""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""64"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Medium Shading 2 Accent 5""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""65"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Medium List 1 Accent 5""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""66"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Medium List 2 Accent 5""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""67"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Medium Grid 1 Accent 5""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""68"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Medium Grid 2 Accent 5""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""69"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Medium Grid 3 Accent 5""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""70"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Dark List Accent 5""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""71"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Colorful Shading Accent 5""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""72"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Colorful List Accent 5""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""73"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Colorful Grid Accent 5""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""60"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Light Shading Accent 6""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""61"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Light List Accent 6""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""62"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Light Grid Accent 6""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""63"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Medium Shading 1 Accent 6""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""64"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Medium Shading 2 Accent 6""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""65"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Medium List 1 Accent 6""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""66"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Medium List 2 Accent 6""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""67"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Medium Grid 1 Accent 6""""/> "
    mailItem.HTMLBody = mailItem.HTMLBody _
        & "<w:LsdException Locked=""""false"""" Priority=""""68"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Medium Grid 2 Accent 6""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""69"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Medium Grid 3 Accent 6""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""70"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Dark List Accent 6""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""71"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Colorful Shading Accent 6""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""72"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Colorful List Accent 6""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""73"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" Name=""""Colorful Grid Accent 6""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""19"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" QFormat=""""true"""" Name=""""Subtle Emphasis""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""21"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" QFormat=""""true"""" Name=""""Intense Emphasis""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""31"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" QFormat=""""true"""" Name=""""Subtle Reference""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""32"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" QFormat=""""true"""" Name=""""Intense Reference""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""33"""" SemiHidden=""""false"""" UnhideWhenUsed=""""false"""" QFormat=""""true"""" Name=""""Book Title""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""37"""" Name=""""Bibliography""""/>" _
        & "<w:LsdException Locked=""""false"""" Priority=""""39"""" QFormat=""""true"""" Name=""""TOC Heading""""/>" _
        & "</w:LatentStyles>" _
        & "</xml><![endif]--><style><!--" _
        & "/* Font Definitions */" _
        & "@font-face" _
        & "    {font-family:Times New Roman;" _
        & "    panose-1:2 15 5 2 2 2 4 3 2 4;" _
        & "    mso-font-charset:0;" _
        & "    mso-generic-font-family:swiss;" _
        & "    mso-font-pitch:variable;" _
        & "    mso-font-signature:-536870145 1073786111 1 0 415 0;}" _
        & "/* Style Definitions */ "
    mailItem.HTMLBody = mailItem.HTMLBody _
        & "p.MsoNormal , li.MsoNormal, div.MsoNormal" _
        & "    {mso-style-unhide:no;" _
        & "    mso-style-qformat:yes;" _
        & "    mso-style-parent:"""""""";" _
        & "    margin:0in;" _
        & "    margin-bottom:.0001pt;" _
        & "    mso-pagination:widow-orphan;" _
        & "    font-size:12.0pt;" _
        & "    font-family:""""Times New Roman"""",""""sans-serif"""";" _
        & "    mso-ascii-font-family:Times New Roman;" _
        & "    mso-fareast-font-family:Times New Roman;" _
        & "    mso-hansi-font-family:Times New Roman;" _
        & "    mso-bidi-font-family:""""Times New Roman"""";}" _
        & "a: link , Span.MsoHyperlink" _
        & "    {mso-style-priority:99;" _
        & "    color:blue;" _
        & "    text-decoration:underline;" _
        & "    text-underline:single;}" _
        & "a: visited , Span.MsoHyperlinkFollowed" _
        & "    {mso-style-noshow:yes;" _
        & "    mso-style-priority:99;" _
        & "    color:purple;" _
        & "    text-decoration:underline;" _
        & "    text-underline:single;} "
    mailItem.HTMLBody = mailItem.HTMLBody _
        & "p.MsoPlainText , li.MsoPlainText, div.MsoPlainText" _
        & "    {mso-style-priority:99;" _
        & "    mso-style-link:""""Plain Text Char"""";" _
        & "    margin:0in;" _
        & "    margin-bottom:.0001pt;" _
        & "    mso-pagination:widow-orphan;" _
        & "    font-size:12.0pt;" _
        & "    mso-bidi-font-size:10.5pt;" _
        & "    font-family:""""Times New Roman"""",""""sans-serif"""";" _
        & "    mso-fareast-font-family:Calibri;" _
        & "    mso-bidi-font-family:""""Times New Roman"""";}" _
        & "Span.EmailStyle17" _
        & "    {mso-style-type:personal-compose;" _
        & "    mso-style-noshow:yes;" _
        & "    mso-style-unhide:no;" _
        & "    mso-ansi-font-size:12.0pt;" _
        & "    mso-bidi-font-size:12.0pt;" _
        & "    font-family:""""Times New Roman"""",""""sans-serif"""";" _
        & "    mso-ascii-font-family:Times New Roman;" _
        & "    mso-hansi-font-family:Times New Roman;" _
        & "    mso-bidi-font-family:""""Times New Roman"""";" _
        & "    color:windowtext;" _
        & "    font-weight:normal;" _
        & "    font-style:normal;} "
    mailItem.HTMLBody = mailItem.HTMLBody _
        & "Span.PlainTextChar" _
        & "    {mso-style-name:""""Plain Text Char"""";" _
        & "    mso-style-priority:99;" _
        & "    mso-style-unhide:no;" _
        & "    mso-style-locked:yes;" _
        & "    mso-style-link:""""Plain Text"""";" _
        & "    mso-bidi-font-size:10.5pt;" _
        & "    font-family:""""Calibri"""",""""sans-serif"""";" _
        & "    mso-ascii-font-family:Calibri;" _
        & "    mso-hansi-font-family:Calibri;}" _
        & "Span.SpellE" _
        & "    {mso-style-name:"""""""";" _
        & "    mso-spl-e:yes;}" _
        & ".MsoChpDefault" _
        & "    {mso-style-type:export-only;" _
        & "    mso-default-props:yes;" _
        & "    font-family:""""Times New Roman"""",""""sans-serif"""";" _
        & "    mso-ascii-font-family:Times New Roman;" _
        & "    mso-fareast-font-family:Times New Roman;" _
        & "    mso-hansi-font-family:Times New Roman;" _
        & "    mso-bidi-font-family:""""Times New Roman"""";}" _
        & "@page WordSection1" _
        & "    {size:8.5in 11.0in;" _
        & "    margin:1.0in 1.0in 1.0in 1.0in; "
    mailItem.HTMLBody = mailItem.HTMLBody _
        & "    mso-header-margin:.5in;" _
        & "    mso-footer-margin:.5in;" _
        & "    mso-paper-source:0;}" _
        & "div.WordSection1" _
        & "    {page:WordSection1;}" _
        & "--></style><!--[if gte mso 10]><style>/* Style Definitions */" _
        & "Table.MsoNormalTable" _
        & "    {mso-style-name:""""Table Normal"""";" _
        & "    mso-tstyle-rowband-size:0;" _
        & "    mso-tstyle-colband-size:0;" _
        & "    mso-style-noshow:yes;" _
        & "    mso-style-priority:99;" _
        & "    mso-style-parent:"""""""";" _
        & "    mso-padding-alt:0in 5.4pt 0in 5.4pt;" _
        & "    mso-para-margin:0in;" _
        & "    mso-para-margin-bottom:.0001pt;" _
        & "    mso-pagination:widow-orphan;" _
        & "    font-size:12.0pt;" _
        & "    font-family:""""Times New Roman"""",""""sans-serif"""";" _
        & "    mso-ascii-font-family:Times New Roman;" _
        & "    mso-hansi-font-family:Times New Roman;}" _
        & "Table.MsoTableGrid" _
        & "    {mso-style-name:""""Table Grid"""";" _
        & "    mso-tstyle-rowband-size:0; "
    mailItem.HTMLBody = mailItem.HTMLBody _
        & "    mso-tstyle-colband-size:0;" _
        & "    mso-style-priority:59;" _
        & "    mso-style-unhide:no;" _
        & "    border:solid windowtext 1.0pt;" _
        & "    mso-border-alt:solid windowtext .5pt;" _
        & "    mso-padding-alt:0in 5.4pt 0in 5.4pt;" _
        & "    mso-border-insideh:.5pt solid windowtext;" _
        & "    mso-border-insidev:.5pt solid windowtext;" _
        & "    mso-para-margin:0in;" _
        & "    mso-para-margin-bottom:.0001pt;" _
        & "    mso-pagination:widow-orphan;" _
        & "    font-size:12.0pt;" _
        & "    font-family:""""Times New Roman"""",""""sans-serif"""";" _
        & "    mso-ascii-font-family:Times New Roman;" _
        & "    mso-hansi-font-family:Times New Roman;}" _
        & "</style><![endif]--><!--[if gte mso 9]><xml>" _
        & "<o:shapedefaults v:ext=""""edit"""" spidmax=""""1026"""" />" _
        & "</xml><![endif]--><!--[if gte mso 9]><xml>" _
        & "<o:shapelayout v:ext=""""edit"""">" _
        & "<o:idmap v:ext=""""edit"""" data=""""1"""" />"
    'The actual mail text starts here
    mailItem.HTMLBody = mailItem.HTMLBody _
        & "</o:shapelayout></xml><![endif]--></head><body lang=EN-US link=blue vlink=purple style='tab-interval:.5in'><div class=WordSection1><p class=MsoPlainText>" _
        & "<p class=MsoPlainText>" & emailBody & "<o:p></o:p></p>" _
        & "<table class=MsoTableGrid border=0 cellspacing=0 cellpadding=0 style='border-collapse:collapse;border:none;mso-yfti-tbllook:1184;mso-padding-alt:0in 5.4pt 0in 5.4pt;mso-border-insideh:none;mso-border-insidev:none'><tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes'><td width=319 valign=top style='width:239.4pt;padding:0in 5.4pt 0in 5.4pt'>" _
        & "<p>Employee Name:<o:p></o:p></p></td><td width=319 valign=top style='width:239.4pt;padding:0in 5.4pt 0in 5.4pt'><p align=right style='text-align:right'>" & second_column_header & "<o:p></o:p></p></td></tr>" _
    'This is the start of the list of employees and their course count. The column headings are here ^.
    mailItem.HTMLBody = mailItem.HTMLBody & employeeList _
         & "</table><p><o:p></o:p></p><p><a href=""www.google.com""> " & emailFooter & "</a> <o:p></o:p></p><p class=MsoPlainText><o:p>&nbsp;</o:p></p></div></body></html>"
        ' The lines below are replaced by the employeeList created in the main MassEmail function. This ^ is the URL and this ^ is the displayed text for the hyperlink in the footer of the email. (The HTML was removed from the front of it because it is a local variable pased, not a global.
        '& "<tr style='mso-yfti-irow:1'><td width=319 valign=top style='width:239.4pt;padding:0in 5.4pt 0in 5.4pt'><p class=MsoPlainText><span class=SpellE>Dain,Purba</span> S<o:p></o:p></p></td><td width=319 valign=top style='width:239.4pt;padding:0in 5.4pt 0in 5.4pt'><p class=MsoPlainText align=right style='text-align:right'>1<o:p></o:p></p></td></tr>" _
        '& "<tr style='mso-yfti-irow:2'><td width=319 valign=top style='width:239.4pt;padding:0in 5.4pt 0in 5.4pt'><p class=MsoPlainText><span class=SpellE>Fetzer,John</span> Tyson<o:p></o:p></p></td><td width=319 valign=top style='width:239.4pt;padding:0in 5.4pt 0in 5.4pt'><p class=MsoPlainText align=right style='text-align:right'>2<o:p></o:p></p></td></tr>" _
        '& "<tr style='mso-yfti-irow:3'><td width=319 valign=top style='width:239.4pt;padding:0in 5.4pt 0in 5.4pt'><p class=MsoPlainText>Goebel <span class=SpellE>Jr</span>.,Kenneth August<o:p></o:p></p></td><td width=319 valign=top style='width:239.4pt;padding:0in 5.4pt 0in 5.4pt'><p class=MsoPlainText align=right style='text-align:right'>3*<o:p></o:p></p></td></tr>" _
        '& "<tr style='mso-yfti-irow:4'><td width=319 valign=top style='width:239.4pt;padding:0in 5.4pt 0in 5.4pt'><p class=MsoPlainText><span class=SpellE>Ruiz,Edward</span><o:p></o:p></p></td><td width=319 valign=top style='width:239.4pt;padding:0in 5.4pt 0in 5.4pt'><p class=MsoPlainText align=right style='text-align:right'>3*<o:p></o:p></p></td></tr>" _
        '& "<tr style='mso-yfti-irow:5;mso-yfti-lastrow:yes'><td width=319 valign=top style='width:239.4pt;padding:0in 5.4pt 0in 5.4pt'><p class=MsoPlainText><span class=SpellE>Hatchett,Deanna</span> M<o:p></o:p></p></td><td width=319 valign=top style='width:239.4pt;padding:0in 5.4pt 0in 5.4pt'><p class=MsoPlainText align=right style='text-align:right'>1*<o:p></o:p></p></td></tr></table>" _
        '& "<p><a href=""""www.google.com""""> " & emailFooter & "</a> <o:p></o:p></p><p class=MsoPlainText><o:p>&nbsp;</o:p></p></div></body></html>"
    mailItem.display
    If Check19.Value <> -1 Then
        SendKeys "%s"   'This is a poor workaround for the outlook automatic email warning. Allow access to Outlook files for 1-5 minutes, rather than 1 request at a time.
        'mailItem.Send  'Because of this workaround you can't type/use the computer while the emails send (Currently not an issue since the code runs in a few seconds/minutes.
    End If
    Set mailItem = Nothing
    Cleanup
End Sub


