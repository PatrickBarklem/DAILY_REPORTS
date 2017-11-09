Attribute VB_Name = "DailyReportCreator"
Public noOfIncidents As Integer
Sub JLT_Daily_Report()
Attribute JLT_Daily_Report.VB_ProcData.VB_Invoke_Func = "q\n14"
'
' JLT_Daily_Report Macro
'
' Keyboard Shortcut: Ctrl+l
'
    
    Dim TitleDate As String
    Dim LoopCounter As Integer
    Dim IncidentData As Collection
    Dim NewIncident As Incident
    Dim IncidentCounter As Integer
    Dim TitleHeadings(3) As String
    Dim TitleSizes(3) As Integer
    Dim RunOnceFlag As Boolean
    Dim StoreSender As String
    
   
    TitleHeadings(0) = "Incident ID"
    TitleHeadings(1) = "Policy"
    TitleHeadings(2) = "Sender"
    TitleHeadings(3) = "Recipients"
    
    TitleSizes(0) = 20.71
    TitleSizes(1) = 41.71
    TitleSizes(2) = 36.86
    TitleSizes(3) = 35
    
    
    IncidentAmountForm.Show

    Worksheets("Sheet1").Activate
    Range("A4").Select
    
    IncidentCounter = 0
    Set IncidentData = New Collection
    
    Do While IncidentCounter <> noOfIncidents
        
        If ActiveCell.Interior.ColorIndex = xlNone Then
            ActiveCell.Offset(1, 0).Select
        Else
            StoreSender = ActiveCell.Value
            ActiveCell.Offset(1, 0).Select
            If ActiveCell.PivotField.Value = "Recipient(s)" Then
                Do
                    If ActiveCell.Interior.ColorIndex <> xlNone And ActiveCell.PivotField.Value = "Recipient(s)" Then
                        Set NewIncident = New Incident
                        NewIncident.Sender = StoreSender
                        NewIncident.Recipient = Replace(ActiveCell.Value, ", dlp@dlp.dlp", "")
                        NewIncident.Recipient = Replace(NewIncident.Recipient, ", ", Chr(10))
                        ActiveCell.Offset(1, 0).Select
                        NewIncident.Policy = ActiveCell.Value
                        ActiveCell.Offset(1, 0).Select
                        NewIncident.ID = ActiveCell.Value
                        IncidentData.Add NewIncident
                        Debug.Print IncidentData(IncidentCounter + 1).ID
                        Debug.Print IncidentData(IncidentCounter + 1).Policy
                        Debug.Print IncidentData(IncidentCounter + 1).Sender
                        Debug.Print IncidentData(IncidentCounter + 1).Recipient
                        ActiveCell.Offset(1, 0).Select
                        IncidentCounter = IncidentCounter + 1
                    Else
                        ActiveCell.Offset(1, 0).Select
                    End If
                Loop Until ActiveCell.PivotField.Value = "Sender"
            End If
        End If
    Loop
        
        
        
        
    With ThisWorkbook
        .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = "Report Table"
    End With
    
    
    If Weekday(Date) <> 2 Then
        TitleDate = Format(Date - 1, "dd-mm-yyyy")
    Else
        TitleDate = Format(Date - 3, "dd-mm-yyyy") & " - " & Format(Date - 1, "dd-mm-yyyy")
    End If
    
    With Range("B2:E2")
        .Merge
        .Select
    End With
    Call Thick_Border
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.349986266670736
        .PatternTintAndShade = 0
    End With
    
    ActiveCell.FormulaR1C1 = "Personal Address Report for " & TitleDate
    With Selection.Font
        .Name = "Calibri"
        .Size = 16
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Bold = True
        .Underline = xlUnderlineStyleSingle
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    
    With Range("B3:E3")
        .Merge
        .Select
    End With
    Call Thick_Border
    Range("B4").Select
    For LoopCounter = 1 To 4
        ActiveCell.FormulaR1C1 = TitleHeadings(LoopCounter - 1)
        ActiveCell.ColumnWidth = TitleSizes(LoopCounter - 1)
        ActiveCell.RowHeight = 22.5
        With Selection
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .WrapText = True
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlContext
                .MergeCells = False
            End With
            With Selection.Font
                .Name = "Calibri"
                .Size = 11
                .Strikethrough = False
                .Superscript = False
                .Subscript = False
                .OutlineFont = False
                .Shadow = False
                .Bold = True
                .Underline = xlUnderlineStyleNone
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0
                .ThemeFont = xlThemeFontNone
            End With
        Call Thick_Border
        ActiveCell.Offset(0, 1).Select
    Next LoopCounter
    Range("B5").Select
    Create_Incident_Rows (IncidentData.Count)
    Range("B5").Select
    Call Populate_Rows(IncidentData, IncidentData.Count)
    
    
    
End Sub
Function Thick_Border()

    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
End Function

Function Create_Incident_Rows(Length As Integer)
    Dim CounterName As Integer
    Dim InnerLoopCounter As Integer
    Dim columnLetter As String
    
    For CounterName = 1 To Length
        For InnerLoopCounter = 1 To 4
            With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = -0.149998474074526
                .PatternTintAndShade = 0
            End With
            With Selection
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .WrapText = True
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlContext
                .MergeCells = False
            End With
            With Selection.Font
                .Name = "Calibri"
                .Size = 11
                .Strikethrough = False
                .Superscript = False
                .Subscript = False
                .OutlineFont = False
                .Shadow = False
                .Underline = xlUnderlineStyleNone
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0
                .ThemeFont = xlThemeFontNone
            End With
            If InnerLoopCounter = 3 Then
                With Selection.Font
                    .Color = -2778277
                    .TintAndShade = 0
                End With
            ElseIf InnerLoopCounter = 4 Then
                With Selection.Font
                    .Color = -16776961
                    .TintAndShade = 0
                End With
            End If
            With Selection.Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            Range(Cells(ActiveCell.Row, ActiveCell.Column), Cells(ActiveCell.Row + 1, ActiveCell.Column)).Select
            Call Thick_Border
            ActiveCell.Offset(0, 1).Select
        Next InnerLoopCounter
        ActiveCell.Offset(2, -4).Select
    Next CounterName

End Function

Function Populate_Rows(Data As Object, Length As Integer)
    Dim CounterName As Integer
    
    For CounterName = 1 To Length
        ActiveCell.FormulaR1C1 = Data(CounterName).ID
        ActiveCell.Offset(0, 1).Select
        ActiveCell.FormulaR1C1 = Data(CounterName).Policy
        ActiveCell.Offset(0, 1).Select
        ActiveCell.FormulaR1C1 = Data(CounterName).Sender
        ActiveCell.Offset(0, 1).Select
        ActiveCell.FormulaR1C1 = Data(CounterName).Recipient
        ActiveCell.Offset(2, -3).Select
    Next CounterName

End Function

