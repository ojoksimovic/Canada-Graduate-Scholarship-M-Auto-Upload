Attribute VB_Name = "Module1"
Sub Research_Portal()

Dim browser As Object
Dim Workbook_Name As String
Dim Worksheet_Name As String
Dim Login_Email As String
Dim Login_Password As String
Dim tags As Object
Dim Last_Name As String
Dim First_Name As String
Dim Application_Title As String


Workbook_Name = ActiveWorkbook.Name
Worksheet_Name = ActiveSheet.Name



'Set studentid = Application.InputBox("Enter the range of student numbers. Do not include the header.", _
Type:=8)

Login_Email = Application.InputBox("Enter your login email:", "Research Portal Login", _
Type:=2)

Login_Password = Application.InputBox("Enter your login password:", "Research Portal Login", _
Type:=2)

Set browser = CreateObject("internetexplorer.application")

browser.Navigate "https://portal-portail.nserc-crsng.gc.ca/db-tb/db-tb.aspx"
browser.Visible = True

Application.Wait (Now + TimeValue("0:00:02"))

While browser.Busy
DoEvents
Wend

browser.Document.All("ctl00_cphMainContent_lgnSiteLogin_UserName").Value = Login_Email
browser.Document.All("ctl00_cphMainContent_lgnSiteLogin_Password").Value = Login_Password


browser.Document.All("ctl00_cphMainContent_lgnSiteLogin_btnLogin").Click

While browser.Busy
DoEvents
Wend

Application.Wait (Now + TimeValue("0:00:04"))

browser.Document.All("ctl00_cphMainContent_InstSloDB_rgrdApplications_ctl00_ctl02_ctl03_FilterTextBox_LastName").Value = " "
browser.Document.All("ctl00_cphMainContent_InstSloDB_rgrdApplications_ctl00_ctl02_ctl03_Filter_LastName").Click

For Each xtag In browser.Document.getElementsByClassName("rmText")
If xtag.innerhtml = "Contains" Then
xtag.Click
Exit For
End If
Next

Application.Wait (Now + TimeValue("0:00:03"))

For Each Cell In Workbooks(Workbook_Name).Worksheets(Worksheet_Name).Range("Last_Name")

Cell.Activate

Last_Name = ActiveCell.Value
First_Name = ActiveCell.Offset(0, 1).Value
Application_Title = ActiveCell.Offset(0, 10).Value

While browser.Busy
DoEvents
Wend

browser.Document.All("ctl00_cphMainContent_InstSloDB_rgrdApplications_ctl00_ctl02_ctl03_FilterTextBox_LastName").Value = Last_Name

browser.Document.All("ctl00_cphMainContent_InstSloDB_rgrdApplications_ctl00_ctl02_ctl03_FilterTextBox_FirstName").Value = First_Name
browser.Document.All("ctl00_cphMainContent_InstSloDB_rgrdApplications_ctl00_ctl02_ctl03_Filter_FirstName").Click

For Each xtag In browser.Document.getElementsByClassName("rmText")
If xtag.innerhtml = "Contains" Then
xtag.Click
Exit For
End If
Next

Application.Wait (Now + TimeValue("0:00:04"))

While browser.Busy
DoEvents
Wend


    
Next

End Sub
