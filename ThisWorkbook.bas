' Infor's workbooks call their Login subroutine on open and Logout subroutine before close
' If you have several files open, logging out every time you close one could be very inconvenient;
' this code will check how many workbooks you have open and only log out if it's the last one.
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Dim minimum_open_workbooks As Integer

    If Workbooks(1).Name = "PERSONAL.XLSB" Then ' If there's a PERSONAL.XLSB open then there is a minimum of two workbooks open.
        minimum_open_workbooks = 2
    Else
        minimum_open_workbooks = 1
    End If

    If Workbooks.Count <= minimum_open_workbooks Then ' If this is the last workbook that's closing, then log out.
        Logout
    End If
End Sub
Private Sub Workbook_Open()
    Login
End Sub
