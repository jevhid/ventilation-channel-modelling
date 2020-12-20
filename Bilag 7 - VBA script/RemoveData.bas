Attribute VB_Name = "RemoveData"

Sub Cleanup()

    ThisWorkbook.Activate
    Sheets("Template").Select
    Range("A2").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.ClearContents
    Sheets("Data").Select
    Range("A1").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.ClearContents
    Sheets("Template").Select
    
End Sub
