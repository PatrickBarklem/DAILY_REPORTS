Attribute VB_Name = "DailyReportCreator"
Public GlobalRowCounter As Integer

Sub JLT_Daily_Report()
Attribute JLT_Daily_Report.VB_ProcData.VB_Invoke_Func = "q\n14"
'
' JLT_Daily_Report Macro
'
' Keyboard Shortcut: Ctrl+q
'
    
    Dim TitleDate As String
    Dim LoopCounter As Integer
    Dim IncidentData As Collection
    Dim NewIncident As Incident
    Dim IncidentCounter As Integer: IncidentCounter = 0
    Dim TitleHeadings(3) As String
    Dim TitleSizes(3) As Integer
    Dim StoreSender As String
    Dim TableStart As Integer
    Dim noOfIncidents As Integer: noOfIncidents = 0
    
    GlobalRowCounter = 1
    
    TitleHeadings(0) = "Incident ID"
    TitleHeadings(1) = "Policy"
    TitleHeadings(2) = "Sender"
    TitleHeadings(3) = "Recipients"
    
    TitleSizes(0) = 20.71
    TitleSizes(1) = 41.71
    TitleSizes(2) = 36.86
    TitleSizes(3) = 35
    
    

    Worksheets("Sheet1").Activate   'Selects the sheet that contains the pivot table and goes to the first incident

    Call Current_Cell
        
    Do
        Call inc(GlobalRowCounter)
    Loop Until Check_If_At_Table(Current_Cell) = True
    
    TableStart = GlobalRowCounter
    
    Do While IsEmpty(Current_Cell) = False
        If InStr(Current_Cell.Value, "@") > 0 Then
            If Current_Cell.PivotField.Value = "Recipient(s)" Then
                Do
                    If Current_Cell.Interior.ColorIndex <> xlNone And Current_Cell.PivotField.Value = "Recipient(s)" Then
                        Call inc(noOfIncidents)
                    End If
                    Call inc(GlobalRowCounter)
                Loop Until Current_Cell.PivotField.Value = "Sender"
            End If
        End If
        Call inc(GlobalRowCounter)
    Loop
    
    
    Set IncidentData = New Collection
    GlobalRowCounter = TableStart
    
    Do While IncidentCounter <> noOfIncidents
            If Current_Cell.Interior.ColorIndex <> xlNone Then StoreSender = Current_Cell.Value
            Call inc(GlobalRowCounter)
            If Current_Cell.PivotField.Value = "Recipient(s)" Then
                Do
                    If Current_Cell.Interior.ColorIndex <> xlNone And Current_Cell.PivotField.Value = "Recipient(s)" Then
                        Set NewIncident = New Incident
                        NewIncident.Sender = StoreSender
                        NewIncident.Recipient = Replace(Current_Cell.Value, ", dlp@dlp.dlp", "")
                        NewIncident.Recipient = Replace(NewIncident.Recipient, ", ", Chr(10))
                        For LoopCounter = 1 To 2
                            Select Case Current_Cell.Offset(LoopCounter, 0).PivotField
                                Case "Policy"
                                    NewIncident.Policy = Current_Cell.Offset(LoopCounter, 0).Value
                                Case "ID"
                                    NewIncident.ID = Current_Cell.Offset(LoopCounter, 0).Value
                            End Select
                        Next LoopCounter
                        IncidentData.Add NewIncident
                        Call inc(IncidentCounter)
                    End If
                    Call inc(GlobalRowCounter)
                Loop Until Current_Cell.PivotField.Value = "Sender"
            End If
        'End If
    Loop
           
    With ThisWorkbook
        .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = "Report Table"
    End With
     
    With Range("B2:E2")
        .Merge
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Pattern = xlSolid
        .FormulaR1C1 = "Personal Address Report for " & Get_Title_Date()
        .Interior.TintAndShade = -0.349986266670736
        .Font.Name = "Calibri"
        .Font.Size = 16
        .Font.Bold = True
        .Font.Underline = xlUnderlineStyleSingle
    End With
    
    Range("B3:E3").Merge
    
    Call Thick_Border(Range("B2:E3"), 1)
    
    With Range("B4:E4")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .RowHeight = 22.5
        .Font.Name = "Calibri"
        .Font.Size = 11
        .Font.Bold = True
    End With
    Call Thick_Border(Range("B4:E4"), 1)
    For LoopCounter = 1 To 4
        With Cells(4, LoopCounter + 1)
            .FormulaR1C1 = TitleHeadings(LoopCounter - 1)
            .ColumnWidth = TitleSizes(LoopCounter - 1)
            .RowHeight = 22.5
        End With
    Next LoopCounter
    Range("B5").Select
    Create_Incident_Rows (IncidentData.Count)
    Range("B5").Select
    Call Populate_Rows(IncidentData, IncidentData.Count)
    
End Sub
Function Thick_Border(CellRange As Range, Optional a As Integer)
    
    If a = 1 Then
        With CellRange.Borders
            .LineStyle = xlContinuous
            .Weight = xlMedium
        End With
    Else
        With CellRange.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlMedium
        End With
        With CellRange.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlMedium
        End With
        With CellRange.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = xlMedium
        End With
        With CellRange.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlMedium
        End With
    End If
    
End Function
Function Create_Incident_Rows(Length As Integer)
    
    Dim CounterName As Integer
    Dim InnerLoopCounter As Integer
    
    For CounterName = 1 To Length
        With Range(ActiveCell, Cells(ActiveCell.Row, ActiveCell.Column + 3))
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = True
            .Interior.Pattern = xlSolid
            .Interior.TintAndShade = -0.149998474074526
            .Font.Name = "Calibri"
            .Font.Size = 11
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).Weight = xlThin
        End With
        For InnerLoopCounter = 1 To 4
            If InnerLoopCounter = 3 Then
                Selection.Font.Color = -2778277
            ElseIf InnerLoopCounter = 4 Then
                Selection.Font.Color = -16776961
            End If
            Call Thick_Border(Range(Cells(ActiveCell.Row, ActiveCell.Column), Cells(ActiveCell.Row + 1, ActiveCell.Column)))
            ActiveCell.Offset(0, 1).Select
        Next InnerLoopCounter
        ActiveCell.Offset(2, -4).Select
    Next CounterName

End Function
Function Populate_Rows(Data As Object, Length As Integer)
    Dim CounterName As Integer
    Dim RowCounter As Integer
    Dim CurrentRow As Integer
    Dim CurrentColumn As Integer
    
    CurrentColumn = ActiveCell.Column
    RowCounter = ActiveCell.Row
    
    For CounterName = 1 To Length
        Cells(RowCounter, CurrentColumn).FormulaR1C1 = Data(CounterName).ID
        Cells(RowCounter, CurrentColumn + 1).FormulaR1C1 = Data(CounterName).Policy
        Cells(RowCounter, CurrentColumn + 2).FormulaR1C1 = Data(CounterName).Sender
        Cells(RowCounter, CurrentColumn + 3).FormulaR1C1 = Data(CounterName).Recipient
        RowCounter = RowCounter + 2
    Next CounterName

End Function
Function Get_Title_Date() As String
    If Weekday(Date) <> 2 Then
        Get_Title_Date = Format(Date - 1, "dd-mm-yyyy")
    Else
        Get_Title_Date = Format(Date - 3, "dd-mm-yyyy") & " - " & Format(Date - 1, "dd-mm-yyyy")
    End If
End Function
Function Check_If_At_Table(CellRange As Range) As Boolean
    
    On Error Resume Next
    
    If IsEmpty(CellRange.PivotField.Value) = False Then Check_If_At_Table = True
    If Err <> 0 Then Check_If_At_Table = False

End Function
Function inc(Number As Integer)
    Number = Number + 1
End Function
Function Current_Cell() As Range
    Set Current_Cell = Cells(GlobalRowCounter, 1)
End Function
