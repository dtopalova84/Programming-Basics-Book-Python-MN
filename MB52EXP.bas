Attribute VB_Name = "MB52EXP"
Public Sub MB52EXP()
        Dim SAPGUI As Object
        Dim app As SAPFEWSELib.GuiApplication
        Dim conn As SAPFEWSELib.GuiConnection
        Dim session As SAPFEWSELib.GuiSession

    Set SAPGUI = GetObject("SAPGUI")
    If IsObject(SAPGUI) Then
        Set app = SAPGUI.GetScriptingEngine
        If IsObject(app) Then
            Set conn = app.Children(0)
            If IsObject(conn) Then
                Set session = conn.Children(0)
    Dim lr
    Dim sht As Worksheet

    Dim appworkbookcount As Long
    appworkbookcount = Application.Workbooks.Count

    With ThisWorkbook.Sheets("CONTROL")
Path = .Range("B2").Value
KillPath = .Range("B3").Value
LogStep "START", "OK", "Run started"

 End With
  Dim plant As String, sloc As String

With ThisWorkbook.Sheets("RULES")
    plant = Trim$(CStr(.Range("B1").Value))
    sloc = Trim$(CStr(.Range("B2").Value))
End With

sloc = Right$("0000" & sloc, 4)

If plant = "" Then plant = "1000"
If sloc = "" Then sloc = "0107"

session.findById("wnd[0]").resizeWorkingPane 137, 25, False
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nmb52"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtMATNR-LOW").Text = ""
session.findById("wnd[0]/usr/ctxtWERKS-LOW").Text = plant
session.findById("wnd[0]/usr/ctxtLGORT-LOW").Text = sloc
session.findById("wnd[0]/usr/ctxtLGORT-LOW").SetFocus
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/usr/lbl[21,13]").SetFocus
session.findById("wnd[0]/usr/lbl[21,13]").caretPosition = 0
session.findById("wnd[0]").sendVKey 43
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").Text = Path
session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "EXPORT.XLSX"
session.findById("wnd[1]/tbar[0]/btn[0]").press
LogStep "EXPORT", "OK", "SAP export triggered"

LogStep "WAIT_FILE", "OK", "Waiting for EXPORT file..."

Dim t0 As Double
t0 = Timer

Do
    DoEvents
    If Dir(KillPath) <> "" Then Exit Do
        Workbooks("EXPORT.XLSX").Close
    Kill KillPath

    If Timer - t0 > 90 Then
        LogStep "WAIT_FILE", "ERR", "Timeout waiting for EXPORT file"
        Exit Sub
    End If
Loop

LogStep "WAIT_FILE", "OK", "EXPORT file detected"
Dim wbExport As Workbook
Set wbExport = Workbooks.Open(KillPath)

Windows("SAP_Excel_MinStock_PRLocation.xlsm").Activate
    Sheets("RAW_MB52").Select
    On Error Resume Next
    ActiveSheet.ShowAllData
    On Error GoTo 0
    Range("A2:L2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Windows("EXPORT.XLSX").Activate
    Range("A2:L2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Windows("SAP_Excel_MinStock_PRLocation.xlsm").Activate
    Range("A2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
            Columns("C:C").Select

LogStep "IMPORT", "OK", "Data copied to RAW_MB52"
Application.DisplayAlerts = False
wbExport.Close SaveChanges:=False
Application.DisplayAlerts = True
Set wbExport = Nothing
DoEvents
Dim t As Double
t = Timer
Do While Timer - t < 3
    DoEvents
Loop
ThisWorkbook.RefreshAll
DoEvents
Application.Wait Now + TimeValue("0:00:20")
DoEvents
LogStep "RESULT", "OK", "Power Query refreshed"

On Error Resume Next
If Dir(KillPath) <> "" Then Kill KillPath
On Error GoTo 0
 Workbooks("EXPORT.XLSX").Close
    Kill KillPath
Call SendMailIfResultHasRows
     End If
      End If
       End If

  End Sub

  Private Sub LogStep(stepName As String, status As String, msg As String)

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("CONTROL")

    Dim nextRow As Long
    nextRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1

    ws.Cells(nextRow, "A").Value = stepName
    ws.Cells(nextRow, "B").Value = Now
    ws.Cells(nextRow, "C").Value = status
    ws.Cells(nextRow, "D").Value = msg

End Sub

Private Sub SendMailIfResultHasRows()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("RESULT")
    Dim toList As String
Dim ccList As String
Dim mailEnabled As String
Dim subjectText As String

mailEnabled = UCase$(GetCtlValue("Mail_Enabled", "YES"))
If mailEnabled <> "YES" Then Exit Sub

toList = GetCtlValue("Mail_To", "")
subjectText = GetCtlValue("Mail_Subject", "ALERT: ROH under minimum")

If toList = "" Then Exit Sub

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    If lastRow < 2 Then Exit Sub

    Dim lastCol As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    Dim rng As Range
    Set rng = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))

    Dim htmlBody As String
    htmlBody = "<p>Здравейте,</p>" & _
               "<p>Има материали под минимум:</p>" & _
               RangeToHTMLTable(rng) & _
               "<p>Поздрави,<br>ERP</p>"

    SendMail_Outlook _
    toList, _
    subjectText & " (" & Format(Now, "yyyy-mm-dd HH:nn") & ")", _
    htmlBody, _
    ""
End Sub

Private Sub SendMail_Outlook(ByVal toList As String, ByVal subjectText As String, _
                             ByVal htmlBody As String, ByVal attachmentPath As String)

    Dim olApp As Object, mail As Object

    On Error Resume Next
    Set olApp = GetObject(, "Outlook.Application")
    On Error GoTo 0
    If olApp Is Nothing Then Set olApp = CreateObject("Outlook.Application")

    Set mail = olApp.CreateItem(0)

    With mail
        .To = toList
        .Subject = subjectText
        .htmlBody = htmlBody
        If Len(attachmentPath) > 0 Then .Attachments.Add attachmentPath
        .Send
    End With
End Sub

Private Function RangeToHTMLTable(ByVal rng As Range) As String
    Dim r As Long, c As Long
    Dim html As String

    html = "<table>"

    For r = 1 To rng.Rows.Count
        html = html & "<tbody><tr>"
        For c = 1 To rng.Columns.Count
            Dim v As String
            v = CStr(rng.Cells(r, c).Value)

            If r = 1 Then
                html = html & "<th>" & EscapeHTML(v) & "</th>"
            Else
                html = html & "<td>" & EscapeHTML(v) & "</td>"
            End If
        Next c
        html = html & "</tr>"
    Next r

    html = html & "</tbody></table>"
    RangeToHTMLTable = html
End Function

Private Function EscapeHTML(ByVal s As String) As String
    s = Replace(s, "&", "&amp;")
    s = Replace(s, "<", "&lt;")
    s = Replace(s, ">", "&gt;")
    s = Replace(s, """", "&quot;")
    EscapeHTML = s
End Function

Private Function GetCtlValue(ByVal keyName As String, Optional ByVal defaultValue As String = "") As String

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("CONTROL")

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    Dim r As Long
    For r = 1 To lastRow
        If Trim$(CStr(ws.Cells(r, "A").Value)) = keyName Then
            GetCtlValue = Trim$(CStr(ws.Cells(r, "B").Value))
            If GetCtlValue = "" Then GetCtlValue = defaultValue
            Exit Function
        End If
    Next r
    GetCtlValue = defaultValue

End Function

Private Sub CloseWorkbooksByPattern(ByVal pattern As String)
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If wb.Name Like pattern Then
            wb.Close SaveChanges:=False
        End If
    Next wb
End Sub
