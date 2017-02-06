' User reported erratic behavior of ActiveX controls when running programs
'
' The subroutine below can be added to a worksheet and run when the controls are formatted correctly;
' it will read the current attributes and generate code to reset them accordingly. Replace the code in
' the sub with the output in the Debug window, and the resulting sub will reset all objects to their
' correct formats.
'
' Based on http://stackoverflow.com/questions/19385803/how-to-stop-activex-objects-automatically-changing-size-in-office
'
Private Sub FixObjects()
    Dim obj As Object
    Dim objname As String
    For Each obj In Me.OLEObjects
        objname = obj.Name
        Debug.Print "    ' " & objname
        Debug.Print "    Me." + objname + ".Left=" + CStr(obj.Left)
        Debug.Print "    Me." + objname + ".Width=" + CStr(obj.Width)
        Debug.Print "    Me." + objname + ".Height=" + CStr(obj.Height)
        Debug.Print "    Me." + objname + ".Top=" + CStr(obj.Top)
        Debug.Print "    Me.Shapes(""" + objname + """).ScaleHeight 1.25, msoFalse, msoScaleFromTopLeft"
        Debug.Print "    Me.Shapes(""" + objname + """).ScaleHeight 0.8, msoFalse, msoScaleFromTopLeft"
        Debug.Print ""
    Next obj
End Sub
