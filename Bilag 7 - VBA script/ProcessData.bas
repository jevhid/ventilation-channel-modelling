Attribute VB_Name = "ProcessData"
Option Explicit
Public filePath, fileName As String

Sub ProcessData()
    
    On Error GoTo err
    Dim LR As Long
    Dim myRow As Integer
    Dim myColumn, LRNumber1, LCNumber As Integer
    Dim LRLetter, LCLetter, ProcedureVariant As String
    
    filePath = OpenWorkbook()
    Workbooks(GetFilenameFromPath(filePath)).Activate

    SheetSelector
    
    With ActiveSheet                                                    'can be optimized as subprocedures
        .Select
        FindFirstUsedCell myRow, myColumn
        LRLetter = Split(Cells(1, myColumn).Address, "$")(1)
        LCNumber = .Cells(myColumn, Columns.count).End(xlToLeft).Column
        LCLetter = Split(Cells(1, LCNumber).Address, "$")(1)
        LR = .Range(LRLetter & Rows.count).End(xlUp).Row
        Workbooks("Skabelon.xlsm").Worksheets("Data").Cells.Clear
        .Range(LRLetter & myColumn & ":" & LCLetter & LR).Copy _
        Workbooks("Skabelon.xlsm").Worksheets("Data").Range("A1")
    End With
    
    Workbooks("Skabelon.xlsm").Worksheets("Data").Activate
    FindFirstUsedCell myRow, myColumn
    DirectionCheck myRow, myColumn, ProcedureVariant 'sends start cell indexes to sub procedure
    Procedure ProcedureVariant, LCLetter
    
err:
    End
    
    
End Sub
'Sub CopyPaste()
'    With ActiveSheet
'        .Select
'        FindFirstUsedCell myRow, myColumn
'        LRLetter = Split(Cells(1, myColumn).Address, "$")(1)
'        LCNumber = .Cells(myColumn, Columns.count).End(xlToLeft).Column
'        LCLetter = Split(Cells(1, LCNumber).Address, "$")(1)
'        LR = .Range(LRLetter & Rows.count).End(xlUp).Row
'        Workbooks("Skabelon.xlsm").Worksheets("Data").Cells.Clear
'        .Range(LRLetter & myColumn & ":" & LCLetter & LR).Copy _
'        Workbooks("Skabelon.xlsm").Worksheets("Data").Range("A1")
'    End With
'End Sub
Function OpenWorkbook()

    Dim fileName, filePath As String
    
    browseFilePath filePath
    Workbooks.Open (filePath)
    OpenWorkbook = filePath
    
End Function

Function GetFilenameFromPath(ByVal strPath As String) As String
' Returns the rightmost characters of a string upto but not including the rightmost '\'
' e.g. 'c:\winnt\win.ini' returns 'win.ini'

    If Right$(strPath, 1) <> "\" And Len(strPath) > 0 Then
        GetFilenameFromPath = GetFilenameFromPath(Left$(strPath, Len(strPath) - 1)) + Right$(strPath, 1)
    End If
End Function


Sub DirectionCheck(myRow, myColumn, ProcedureVariant)

    Dim x, c, LCLetter, LRLetter, cell As String
    Dim i, r, count  As Integer
    Dim LRNumber, LCNumber As Long

    With Sheets("Data")
        LRLetter = Split(Cells(1, myColumn).Address, "$")(1)
        LRNumber = .Range(LRLetter & Rows.count).End(xlUp).Row
        LCNumber = .Cells(myColumn, Columns.count).End(xlToLeft).Column
        
        For i = 1 To LCNumber
            LCLetter = Split(Cells(1, i).Address, "$")(1)
            x = Range(LCLetter & 1).Value
            If x = "Rumnavne" Or x = "Number" Or x = "Area" Or x = "Room: Department" Then
                count = count + 1
            End If
        Next i
        
        If count >= 2 Then
            ProcedureVariant = "ColumnHeaders"
            'MsgBox "Data  has column headers"
        ElseIf count = 1 Then
            ProcedureVariant = "RowHeaders"
            'MsgBox "Data  has row headers"
        Else
            MsgBox "No data found/Invalid data"
            Exit Sub
        End If
        
    End With
    
End Sub

Sub FindFirstUsedCell(myRow, myColumn)

    Dim rFound As Range

    
    On Error Resume Next
    Set rFound = Cells.Find(What:="*", _
                            After:=Cells(Rows.count, Columns.count), _
                            LookAt:=xlPart, _
                            LookIn:=xlFormulas, _
                            SearchOrder:=xlByRows, _
                            SearchDirection:=xlNext, _
                            MatchCase:=False)
    
    On Error GoTo 0
    
    If rFound Is Nothing Then
        MsgBox "All cells are blank."
    Else
        'MsgBox "First Cell: " & rFound.Address
    End If
    
    rFound.Select
    myRow = ActiveCell.Row
    myColumn = ActiveCell.Column
    
End Sub

Sub PasteRangeColumn(Sourcesheet, Header, DestinationRange, LCLetter)

    Dim xRg As Range
    Dim xRgUni As Range
    Dim xFirstAddress As String
    Dim xStr As String
    Dim newRange As Range
    
    Workbooks("Skabelon.xlsm").Worksheets(Sourcesheet).Activate
    On Error Resume Next
    xStr = Header
    Set xRg = Range("A1:" & LCLetter & "1").Find(xStr, , xlValues, xlWhole, , , True)
    If Not xRg Is Nothing Then
        xFirstAddress = xRg.Address
        Do
            Set xRg = Range("A1:" & LCLetter & "1").FindNext(xRg)
            If xRgUni Is Nothing Then
                Set xRgUni = xRg
            Else
                Set xRgUni = Application.Union(xRgUni, xRg)
            End If
        Loop While (Not xRg Is Nothing) And (xRg.Address <> xFirstAddress)
    End If
    
    xRgUni.Select
    ActiveCell.Offset(1, 0).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("Template").Select
    Range(DestinationRange).Select
    ActiveSheet.Paste

End Sub

Sub Procedure(ProcedureVariant, LCLetter)

    Dim Arr As Variant
    Dim SourceRange, TargetRange As Variant

    If ProcedureVariant = "ColumnHeaders" Then
        PasteRangeColumn "Data", "Rumnavne", "A2", LCLetter
        PasteRangeColumn "Data", "Number", "B2", LCLetter
        PasteRangeColumn "Data", "Specified Supply Airflow", "C2", LCLetter
        PasteRangeColumn "Data", "Specified Return Airflow", "E2", LCLetter
        PasteRangeColumn "Data", "Area", "F2", LCLetter
        PasteRangeColumn "Data", "Room: Department", "G2", LCLetter
    ElseIf ProcedureVariant = "RowHeaders" Then
        SourceRange = ActiveSheet.Range(Selection, Selection.End(xlDown).End(xlToRight))
        TargetRange = "A11"
        Transpose "Skabelon.xlsm", "Data"
        'MsgBox "Inside RowHeaders procedure"
        
        PasteRangeColumn "Data", "Rumnavne", "A2", LCLetter
        PasteRangeColumn "Data", "Number", "B2", LCLetter
        PasteRangeColumn "Data", "Specified Supply Airflow", "C2", LCLetter
        PasteRangeColumn "Data", "Specified Return Airflow", "E2", LCLetter
        PasteRangeColumn "Data", "Area", "F2", LCLetter
        PasteRangeColumn "Data", "Room: Department", "G2", LCLetter
    Else
        MsgBox "Invalid ProcedureVariant"
    End If
    
End Sub

Sub Transpose(Workbook, Sourcesheet)
    
    Workbooks(Workbook).Worksheets(Sourcesheet).Activate
    Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Name = "Transpose"
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
                           False, Transpose:=True
    Sheets(Sourcesheet).Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Sheets("Transpose").Select
    Selection.Cut
    Sheets(Sourcesheet).Select
    Range("A1").Select
    ActiveSheet.Paste
    Sheets("Transpose").Select
    Application.DisplayAlerts = False
    ActiveWindow.SelectedSheets.Delete
    Application.DisplayAlerts = True
    
End Sub

Sub browseFilePath(filePath)
    
    On Error GoTo err
    Dim fileExplorer As FileDialog
    Set fileExplorer = Application.FileDialog(msoFileDialogFilePicker)

    'To allow or disable to multi select
    fileExplorer.AllowMultiSelect = False

    With fileExplorer
        If .Show = -1 Then 'Any file is selected
            filePath = .SelectedItems.Item(1)
        Else ' else dialog is cancelled
            MsgBox "You have cancelled the dialogue"
            filePath = "" ' when cancelled set blank as file path.
            End
        End If
    End With

err:
    Exit Sub
    
End Sub



