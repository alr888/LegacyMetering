Option Explicit
Public Function testarr(ByVal x As String) As String
Dim arr As Variant
Dim retarr As String
  arr = Split(x, ":")
  testarr = arr(1)

End Function

 Sub CopyToExcel(txtBody As String, tDate As Date)
 Dim xlApp As Excel.Application
 Dim xlWB As Object
 Dim xlSheet As Object
 Dim vText1, vText2, vText3, vText4, vText5, vText6, vText7, vText8, vText9, vText10, vText13a, vText13b, vText14 As String
 Dim sText As String
 Dim rCount As Long
 Dim bXStarted As Boolean
 Dim enviro As String
 Dim strPath As String
 
'the path of the workbook
 strPath = "C:\LMG\SERVICE_REQUESTS.xlsx"
     On Error Resume Next
     Set xlApp = New Excel.Application
     If Err <> 0 Then
         Application.StatusBar = "Please wait while Excel source is opened ... "
         Set xlApp = CreateObject("Excel.Application")
         bXStarted = True
     End If
     On Error GoTo 0
     'Open the workbook to input the data
     Set xlWB = xlApp.Workbooks.Open(strPath)
     Set xlSheet = xlWB.Sheets("SERVICE_REQUESTS")
    ' Process the message record
    'Set olItem = Application.ActiveExplorer().Selection(1)arr
    
    'Find the next empty line of the worksheet
     rCount = xlSheet.Range("B" & xlSheet.Rows.Count).End(-4162).Row
     rCount = rCount + 1
     
     
     sText = txtBody  'olItem.Body
     
     Dim sText2 As String
     
     
    Dim xx As String
    
    Dim i, a As Integer
    
    a = 1
    
    For i = 1 To Len(sText)
        xx = Mid(sText, i, 1)
        If xx = Chr(10) Then
            If Mid(sText, i - 1, 1) = Chr(13) Then
                sText2 = (Mid(sText, a, (i - a) - 1))
                If Len(sText2) > 2 Then
                    If InStr(sText2, "Job Number ") > 1 Then
                        vText2 = testarr(sText2)  'Right(sText2, InStr(sText2, ": ") - 1)
                    ElseIf InStr(sText2, "Customer Name ") > 1 Then
                        vText4 = testarr(sText2)
                    ElseIf InStr(sText2, "Address ") > 1 Then
                        vText5 = testarr(sText2)
                    ElseIf InStr(sText2, "Premise / ICP ") > 1 Then
                        vText6 = testarr(sText2)
                    ElseIf InStr(sText2, "Customer Phone ") > 1 Then
                        vText7 = testarr(sText2)
                    ElseIf InStr(sText2, "Date Required ") > 1 Then
                        vText8 = testarr(sText2)
                    ElseIf InStr(sText2, "Dog on Site ") > 1 Then
                        vText9 = testarr(sText2)
                    ElseIf InStr(sText2, "Job Description ") > 1 Then
                        vText10 = testarr(sText2)
                    ElseIf InStr(sText2, "Meter Number ") > 1 Then
                        vText13a = testarr(sText2)
                    ElseIf InStr(sText2, "Additional Meter ") > 1 Then
                        vText13b = testarr(sText2)
                    ElseIf InStr(sText2, "Meter Location ") > 1 Then
                        vText14 = testarr(sText2)
                   
                    Else
                        
                    End If
                End If
                a = i
            End If
        End If
    Next
 
   xlSheet.Range("a" & rCount) = "NOVA"
   xlSheet.Range("b" & rCount) = vText2
   xlSheet.Range("c" & rCount) = DateValue(tDate)
   xlSheet.Range("d" & rCount) = vText4
   xlSheet.Range("e" & rCount) = vText5
   xlSheet.Range("f" & rCount) = vText6
   xlSheet.Range("g" & rCount) = vText7
   xlSheet.Range("h" & rCount) = vText8
   xlSheet.Range("i" & rCount) = vText9
   xlSheet.Range("j" & rCount) = vText10
   xlSheet.Range("k" & rCount) = ""
   xlSheet.Range("l" & rCount) = ""
   xlSheet.Range("m" & rCount) = vText13a & " " & Chr(13) & vText13b
   xlSheet.Range("n" & rCount) = vText14
   xlSheet.Range("o" & rCount) = "DELT"
   xlSheet.Range("w" & rCount) = "READY"
  
     xlWB.Close 1
     If bXStarted Then
         xlApp.Quit
     End If
     Set xlApp = Nothing
     Set xlWB = Nothing
     Set xlSheet = Nothing
 End Sub

Sub CopyTo_LMG_Template(txtBody As String, tDate As Date)
 Dim xlApp As Excel.Application
 Dim xlWB As Object
 Dim xlSheet As Object
 Dim vText1, vText2, vText3, vText4, vText5, vText6, vText7, vText8, vText9, vText10, vText13a, vText13b, vText14 As String
 Dim sText As String
 Dim rCount As Long
 Dim bXStarted As Boolean
 Dim enviro As String
 Dim strPath As String
 
'The path of the LMG WO template
 strPath = "C:\LMG\LMG_WO.xlsx"
     On Error Resume Next
     Set xlApp = New Excel.Application
     If Err <> 0 Then
         Application.StatusBar = "Please wait while Excel source is opened ... "
         Set xlApp = CreateObject("Excel.Application")
         bXStarted = True
     End If
     On Error GoTo 0
     'Open the workbook to input the data
     Set xlWB = xlApp.Workbooks.Open(strPath)
     Set xlSheet = xlWB.Sheets("SERVICE_REQUESTS")
    ' Process the message record
    'Set olItem = Application.ActiveExplorer().Selection(1)arr
    
    sText = txtBody  'olItem.Body
     
     Dim sText2 As String
     
     
    Dim xx As String
    
    Dim i, a As Integer
    
    a = 1
    
    For i = 1 To Len(sText)
        xx = Mid(sText, i, 1)
        If xx = Chr(10) Then
            If Mid(sText, i - 1, 1) = Chr(13) Then
                sText2 = (Mid(sText, a, (i - a) - 1))
                If Len(sText2) > 2 Then
                    If InStr(sText2, "Job Number ") > 1 Then
                        vText2 = testarr(sText2)  'Right(sText2, InStr(sText2, ": ") - 1)
                    ElseIf InStr(sText2, "Customer Name ") > 1 Then
                        vText4 = testarr(sText2)
                    ElseIf InStr(sText2, "Address ") > 1 Then
                        vText5 = testarr(sText2)
                    ElseIf InStr(sText2, "Premise / ICP ") > 1 Then
                        vText6 = testarr(sText2)
                    ElseIf InStr(sText2, "Customer Phone ") > 1 Then
                        vText7 = testarr(sText2)
                    ElseIf InStr(sText2, "Date Required ") > 1 Then
                        vText8 = testarr(sText2)
                    ElseIf InStr(sText2, "Dog on Site ") > 1 Then
                        vText9 = testarr(sText2)
                    ElseIf InStr(sText2, "Job Description ") > 1 Then
                        vText10 = testarr(sText2)
                    ElseIf InStr(sText2, "Meter Number ") > 1 Then
                        vText13a = testarr(sText2)
                    ElseIf InStr(sText2, "Additional Meter ") > 1 Then
                        vText13b = testarr(sText2)
                    ElseIf InStr(sText2, "Meter Location ") > 1 Then
                        vText14 = testarr(sText2)
                    Else
                        
                    End If
                End If
                a = i
            End If
        End If
    Next
 
   xlSheet.Range("b2") = "NOVA"
   xlSheet.Range("b3") = vText2
   xlSheet.Range("b4") = DateValue(tDate)
   xlSheet.Range("b5") = vText4
   xlSheet.Range("b6") = vText5
   xlSheet.Range("b7") = vText6
   xlSheet.Range("b8") = vText7
   xlSheet.Range("b11") = vText8
   xlSheet.Range("b12") = vText9
   xlSheet.Range("b13") = vText10
   xlSheet.Range("b18") = vText13a & " " & Chr(13) & vText13b
   xlSheet.Range("b19") = vText14
   xlSheet.Range("b21") = "DELT"
  
    'Save C:\LMG\LMG_WO.xlsx file as LMG_WO_JobNo_yyyymmdd.xlsx
     xlWB.SaveAs "C:\LMG\LMG_WO_" & vText2 & "_" & Format(Date, "yyyymmdd") & ".xlsx"
     xlWB.Close 1
     If bXStarted Then
         xlApp.Quit
     End If
     Set xlApp = Nothing
     Set xlWB = Nothing
     Set xlSheet = Nothing
 End Sub


Sub Parse_SR_Inbox()

Dim myOlApp As New Outlook.Application
Dim objNamespace As Outlook.NameSpace
Dim objFolder As Outlook.MAPIFolder
Dim filteredItems As Outlook.Items
Dim itm As Object
Dim Found As Boolean
Dim strFilter, tDate As String


Set objNamespace = myOlApp.GetNamespace("MAPI")
Set objFolder = objNamespace.GetDefaultFolder(olFolderInbox)

strFilter = "@SQL=" & Chr(34) & "urn:schemas:httpmail:subject" & Chr(34) & " like '%Service Request Case%'"

Set filteredItems = objFolder.Items.Restrict(strFilter)

If filteredItems.Count = 0 Then
    Debug.Print "No emails found"
    Found = False
Else
    Found = True
    ' this loop is optional, it displays the list of emails by subject.
    For Each itm In filteredItems
     'Debug.Print itm.Subject
     'tDate = DateValue(itm.CreationTime)
      Call CopyToExcel(itm.Body, itm.CreationTime)
     'MsgBox itm.Subject
    Next
End If


'If the subject isn't found:
If Not Found Then
    'NoResults.Show
Else
   Debug.Print "Found " & filteredItems.Count & " items."

End If

'myOlApp.Quit
Set myOlApp = Nothing

MsgBox "Parsing completed review data in C:\LMG\SERVICE_REQUESTS.xlsx", vbInformation



End Sub

Sub Parse_SR_To_LMG_Template()

Dim myOlApp As New Outlook.Application
Dim objNamespace As Outlook.NameSpace
Dim objFolder As Outlook.MAPIFolder
Dim filteredItems As Outlook.Items
Dim itm As Object
Dim Found As Boolean
Dim strFilter, tDate As String


Set objNamespace = myOlApp.GetNamespace("MAPI")
Set objFolder = objNamespace.GetDefaultFolder(olFolderInbox)

strFilter = "@SQL=" & Chr(34) & "urn:schemas:httpmail:subject" & Chr(34) & " like '%Service Request Case%'"

Set filteredItems = objFolder.Items.Restrict(strFilter)

If filteredItems.Count = 0 Then
    Debug.Print "No emails found"
    Found = False
Else
    Found = True
    'MsgBox "Meron"
    ' this loop is optional, it displays the list of emails by subject.
    For Each itm In filteredItems
     'Debug.Print itm.Subject
     'tDate = DateValue(itm.CreationTime)
      Call CopyTo_LMG_Template(itm.Body, itm.CreationTime)
     'MsgBox itm.Subject
    Next
End If


'If the subject isn't found:
If Not Found Then
    'MsgBox "Wala"
    'NoResults.Show
Else
   Debug.Print "Found " & filteredItems.Count & " items."

End If

'myOlApp.Quit
Set myOlApp = Nothing

MsgBox "Parsing completed check LMG template.", vbInformation



End Sub

