Sub SendMail_1()
'this is dog-shit, improvements are needed to recursively call the same mail object, whilst incrementing the rows. For each new mail you have to create a new function!!
    Dim Wks    As Worksheet
    Dim OutMail As Object
    Dim OutApp As Object
    Dim myRng  As Range
    Dim list   As Object
    Dim item   As Variant
    Dim LastRow As Long
    Dim uniquesArray()
    Dim Dest   As String
    Dim strbody
    
    Set list = CreateObject("System.Collections.ArrayList")
    Set Wks = ThisWorkbook.Sheets("Sheet1")
    
    With Application
        .EnableEvents = False
        .ScreenUpdating = False
    End With
    
    With Wks
        For Each item In .Range("A4", .Range("A" & .Rows.Count).End(xlUp))
            'If Not list.Contains(item.value)Then
            list.Add item.Value
        Next
    End With
    
    For Each item In list
        
        Wks.Range("B1:H" & Range("A" & Rows.Count).End(xlUp).Row).AutoFilter Field:=1, Criteria1:=1
        Set OutApp = CreateObject("Outlook.Application")
        Set OutMail = OutApp.CreateItem(0)
        On Error Resume Next
        
        LastRow = Cells(Rows.Count, "A").End(xlUp).Row
        Set myRng = Wks.Range("B1:G" & LastRow).SpecialCells(xlCellTypeVisible)
        
        Dest = Cells(LastRow, "H").Value
        strbody = "Dear ," & "<br>" & _
                  "These are your total open rejects: " & "<br/><br>"
        
        With OutMail
            .To = Dest
            .CC = ""
            .BCC = ""
            .Subject = "Weekly totals"
            .HTMLBody = strbody & RangetoHTML(myRng)
            .Display
            '.Send
        End With
        On Error GoTo 0
    Next
    
    On Error Resume Next
    Wks.ShowAllData
    On Error GoTo 0
    
    With Application
        .EnableEvents = True
        .ScreenUpdating = True
    End With
    
End Sub
Sub SendMail_2()

    Dim Wks    As Worksheet
    Dim OutMail As Object
    Dim OutApp As Object
    Dim myRng  As Range
    Dim list   As Object
    Dim item   As Variant
    Dim LastRow As Long
    Dim uniquesArray()
    Dim Dest   As String
    Dim strbody
    
    Set list = CreateObject("System.Collections.ArrayList")
    Set Wks = ThisWorkbook.Sheets("Sheet1")
    
    With Application
        .EnableEvents = False
        .ScreenUpdating = False
    End With
    
    With Wks
        For Each item In .Range("A4", .Range("A" & .Rows.Count).End(xlUp))
            If Not list.Contains(item.Value) Then list.Add item.Value
        Next
    End With

    For Each item In list
        
        Wks.Range("A1:H" & Range("A" & Rows.Count).End(xlUp).Row).AutoFilter Field:=1, Criteria1:=2
        Set OutApp = CreateObject("Outlook.Application")
        Set OutMail = OutApp.CreateItem(0)
        On Error Resume Next
        
        LastRow = Cells(Rows.Count, "A").End(xlUp).Row
        Set myRng = Wks.Range("B1:G" & LastRow).SpecialCells(xlCellTypeVisible)
        
        Dest = Cells(LastRow, "H").Value
        strbody = "Dear ," & "<br>" & _
                  "These are your total open rejects: " & "<br/><br>"
        
        With OutMail
            .To = Dest
            .CC = ""
            .BCC = ""
            .Subject = "Weekly totals"
            .HTMLBody = strbody & RangetoHTML(myRng)
            .Display
            '.Send
        End With
        On Error GoTo 0
    Next
    
    On Error Resume Next
    Wks.ShowAllData
    On Error GoTo 0
    
    With Application
        .EnableEvents = True
        .ScreenUpdating = True
    End With
    
End Sub
Sub SendMail_3()

    Dim Wks    As Worksheet
    Dim OutMail As Object
    Dim OutApp As Object
    Dim myRng  As Range
    Dim list   As Object
    Dim item   As Variant
    Dim LastRow As Long
    Dim uniquesArray()
    Dim Dest   As String
    Dim strbody
    
    Set list = CreateObject("System.Collections.ArrayList")
    Set Wks = ThisWorkbook.Sheets("Sheet1")
    
    With Application
        .EnableEvents = False
        .ScreenUpdating = False
    End With
    
    With Wks
        For Each item In .Range("A4", .Range("A" & .Rows.Count).End(xlUp))
            If Not list.Contains(item.Value) Then list.Add item.Value
        Next
    End With
    
    For Each item In list
        
        Wks.Range("A1:H" & Range("A" & Rows.Count).End(xlUp).Row).AutoFilter Field:=1, Criteria1:=3
        Set OutApp = CreateObject("Outlook.Application")
        Set OutMail = OutApp.CreateItem(0)
        On Error Resume Next
        
        LastRow = Cells(Rows.Count, "A").End(xlUp).Row
        Set myRng = Wks.Range("B1:G" & LastRow).SpecialCells(xlCellTypeVisible)
        
        Dest = Cells(LastRow, "H").Value
        strbody = "Dear ," & "<br>" & _
                  "These are your total open rejects: " & "<br/><br>"
        
        With OutMail
            .To = Dest
            .CC = ""
            .BCC = ""
            .Subject = "Weekly totals"
            .HTMLBody = strbody & RangetoHTML(myRng)
            .Display
            '.Send
        End With
        On Error GoTo 0
    Next
    
    On Error Resume Next
    Wks.ShowAllData
    On Error GoTo 0
    
    With Application
        .EnableEvents = True
        .ScreenUpdating = True
    End With
    
End Sub
Function RangetoHTML(myRng As Range)

    Dim TempFile As String
    Dim TempWB As Workbook
    Dim fso    As Object
    Dim ts     As Object
    Dim i      As Integer

    TempFile = Environ$("temp") & "\" & Format(Now, "dd-mm-yy h-mm-ss") & ".htm"
    myRng.Copy
    Set TempWB = Workbooks.Add(1)
    With TempWB.Sheets(1)
        .Cells(1).PasteSpecial Paste:=8
        .Cells(1).PasteSpecial xlPasteAllUsingSourceTheme, , False, False
        .Cells(1).Select
        Application.CutCopyMode = False
        On Error Resume Next
        .DrawingObjects.Visible = True
        .DrawingObjects.Delete
        On Error GoTo 0
        
        For i = 7 To 12
            With .UsedRange.Borders(i)
                .LineStyle = xlContinuous
                .ColorIndex = xlAutomatic
                .TintAndShade = 0
                .Weight = xlMedium
            End With
        Next i
    End With
    
    With TempWB.PublishObjects.Add( _
         SourceType:=xlSourceRange, _
         Filename:=TempFile, _
         Sheet:=TempWB.Sheets(1).Name, _
         Source:=TempWB.Sheets(1).UsedRange.Address, _
         HtmlType:=xlHtmlStatic)
        .Publish (True)
    End With
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(TempFile).OpenAsTextStream(1, -2)
    RangetoHTML = ts.readall
    ts.Close
    RangetoHTML = Replace(RangetoHTML, "align=center x:publishsource=", _
                  "align=left x:publishsource=")
    TempWB.Close savechanges:=False
    Kill TempFile
    Set ts = Nothing
    Set fso = Nothing
    Set TempWB = Nothing
End Function



Sub CallMacros()

Call SendMail_1
Call SendMail_2
Call SendMail_3

End Sub
