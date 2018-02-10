Attribute VB_Name = "EncAuto"
Sub EncA()
On Error GoTo erroc

Dim PKq As String, opt As Variant, Boundary As Range
PKq = InputBox("Set Password (0-9A-Za-z!@#!%...)")

Dim rng As Range

If Selection.Cells.Count > 1 Then
    opt = Msgbox("Do you want to encrypt only selected cells?", vbYesNoCancel)
    If opt = vbCancel Then Exit Sub
    If opt = vbYes Then
        Set Boundary = Selection
    Else
        Set Boundary = ActiveSheet.UsedRange
    End If
Else
    Set Boundary = ActiveSheet.UsedRange
End If

On Error Resume Next
For Each rng In Boundary

    If rng.Value = vbNullString Then
    Else
    rng.Value = Encode(CStr(rng.Value), PKq)
    End If

Next

Application.ScreenUpdating = True: Application.EnableEvents = True: Application.Calculation = xlCalulationAutomatic
Msgbox "Complete by PKey:" + PKq
Exit Sub

erroc:
Msgbox "Error"
Application.ScreenUpdating = True: Application.EnableEvents = True: Application.Calculation = xlCalulationAutomatic

End Sub

Sub DecA()
Application.ScreenUpdating = False: Application.EnableEvents = False: Application.Calculation = xlCalculationManual

On Error GoTo erroc

Dim PKq As String, opt As Variant, Boundary As Range
PKq = InputBox("Password (0-9A-Za-z!@#!%...)")

Dim rng As Range

If Selection.Cells.Count > 1 Then
    opt = Msgbox("Do you want to decrypt only selected cells?", vbYesNoCancel)
    If opt = vbCancel Then Exit Sub
    If opt = vbYes Then
        Set Boundary = Selection
    Else
        Set Boundary = ActiveSheet.UsedRange
    End If
Else
    Set Boundary = ActiveSheet.UsedRange
End If

On Error Resume Next

For Each rng In Boundary
    If rng.Value = vbNullString Then
    Else
    rng.Value = Decode(CStr(rng.Value), PKq)
    End If
Next

Application.ScreenUpdating = True: Application.EnableEvents = True: Application.Calculation = xlCalulationAutomatic
Msgbox "Complete by PKey:" + PKq
Exit Sub

erroc:
Msgbox "Error"
Application.ScreenUpdating = True: Application.EnableEvents = True: Application.Calculation = xlCalulationAutomatic

End Sub
