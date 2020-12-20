Attribute VB_Name = "SheetSelect"

Sub SheetSelector()

Const ColItems  As Long = 20
Const LetterWidth As Long = 20
Const HeightRowz As Long = 18
Const SheetID As String = "__SheetSelection"
 
Dim i%, TopPos%, iSet%, optCols%, intLetters%, optMaxChars%, optLeft%
Dim wsDlg As DialogSheet, objOpt As OptionButton, optCaption$, objSheet As Object
optCaption = "": i = 0
 
Application.ScreenUpdating = False
 
On Error Resume Next
Application.DisplayAlerts = False
ActiveWorkbook.DialogSheets(SheetID).Delete
Application.DisplayAlerts = True
err.Clear
 
Set wsDlg = ActiveWorkbook.DialogSheets.Add
With wsDlg
.Name = SheetID
.Visible = xlSheetHidden
iSet = 0: optCols = 0: optMaxChars = 0: optLeft = 78: TopPos = 40
 
For Each objSheet In ActiveWorkbook.Sheets
If objSheet.Visible = xlSheetVisible Then
i = i + 1
 
If i Mod ColItems = 1 Then
optCols = optCols + 1
TopPos = 40
optLeft = optLeft + (optMaxChars * LetterWidth)
optMaxChars = 0
End If
 
intLetters = Len(objSheet.Name)
If intLetters > optMaxChars Then optMaxChars = intLetters
iSet = iSet + 1
.OptionButtons.Add optLeft, TopPos, intLetters * LetterWidth, 16.5
.OptionButtons(iSet).Text = objSheet.Name
TopPos = TopPos + 13
 
End If
Next objSheet
 
If i > 0 Then
 
.Buttons.Left = optLeft + (optMaxChars * LetterWidth) + 24
 
With .DialogFrame
.Height = Application.Max(68, WorksheetFunction.Min(iSet, ColItems) * HeightRowz + 10)
.Width = optLeft + (optMaxChars * LetterWidth) + 24
.Caption = "Select sheet to go to"
End With
 
.Buttons("Button 2").BringToFront
.Buttons("Button 3").BringToFront
Application.ScreenUpdating = True
 
If .Show = True Then
For Each objOpt In wsDlg.OptionButtons
If objOpt.Value = xlOn Then
optCaption = objOpt.Caption
Exit For
End If
Next objOpt
End If
 
If optCaption = "" Then
MsgBox "You did not select a worksheet.", 48, "Cannot continue"
End
Else
 
'MsgBox "You selected the sheet named ''" & optCaption & "''." & vbCrLf & "Click OK to go there.", 64, "FYI:"
Sheets(optCaption).Activate
 
End If
 
End If
 
Application.DisplayAlerts = False
.Delete
Application.DisplayAlerts = True
 
End With

End Sub
