' Lawson Journal Entry Tool
' Copyright (C) 2016 Joe Carey
'
' This program is free software: you can redistribute it and/or modify it under the terms of the GNU General
' Public License as published by the Free Software Foundation, either version 3 of the License, or (at your
' option) any later version.
'
' This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the
' implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License
' for more details.
'
' You should have received a copy of the GNU General Public License along with this program. If not, see
' <http://www.gnu.org/licenses/>.
'
' Home is https://github.com/indepndnt/vba_Lawson10_JournalEntry_with_Query/
'
Option Explicit
Public Sub inBalance()
On Error GoTo errHandler
    Dim s As String
    Dim rngOut As Range ' Output table range

    If Not CheckUserAttributes() Then Login
    Set rngOut = Me.Range("query_output")(2, 1)
    Call fDeleteFrom(rngOut) ' delete any previous outputs and reset UsedRange
    Range(Me.Range("query_output")(2, 1), Me.Range("query_output")(2, 19)).Clear ' clear first row without deleting table formula
    Me.Range("query_errors").EntireRow.Clear
    Me.Range("query_errors").Value = "Error messages go here:"

    s = "PROD=" & g_sProductLine ' Start building POST data string with Product Line
    s = s & "&FILE=GLAMOUNTS&INDEX=GAMSET1" ' Table GLAMOUNTS, criteria set GAMSET1: key = co=fy=acct-unit=account
    s = s & "&KEY=" & FilterForWeb(Me.Range("query_company").Value & "=" & Me.Range("query_fy").Value & "=" & _
        Me.Range("query_acctunit").Value & "=" & Me.Range("query_account").Value)
    s = s & "&FIELD=" & FilterForWeb("COMPANY;ACCT-UNIT;ACCOUNT;SUB-ACCOUNT;CHART-DETAIL.ACCOUNT-DESC;FISCAL-YEAR;CYBAMT;CYPAMT1;CYPAMT2;CYPAMT3;CYPAMT4;CYPAMT5;CYPAMT6;CYPAMT7;CYPAMT8;CYPAMT9;CYPAMT10;CYPAMT11;CYPAMT12")
    s = s & "&OUT=XML&NEXT=FALSE" ' NEXT=FALSE means don't give me the RELOAD string.
    s = s & "&MAX=10000&keyUsage=PARAM" ' Give me up to 600 records.
    s = SendURL(s, "D")
    SetXMLObject ' Load IE page document into XML document object
    If Not g_oDom.LoadXML(s) Then
        Me.Cells(Me.Range("query_errors").row, Me.Cells(Me.Range("query_errors").row, 16384).End(xlToLeft).Column + 1).Value = "Could not load XML data from server."
        Exit Sub
    End If
    If Not inXmlDme Then ' do we have a /DME xml document?
        Me.Cells(Me.Range("query_errors").row, Me.Cells(Me.Range("query_errors").row, 16384).End(xlToLeft).Column + 1).Value = inXmlData("/ERROR/MSG", 1) & " GLAMOUNTS Query" ' Error message from GLAMOUNTS query
        Exit Sub
    End If

    Application.ScreenUpdating = False ' Faster w/o updating screen for each cell - especially if output range is a table
    Dim qArray() As String
    Call inQueryArray(qArray)
    Call inArrayToRange(qArray, rngOut)
    Application.ScreenUpdating = True ' Revert to showing user the output

    Exit Sub
errHandler:
    Me.Cells(Me.Range("query_errors").row, Me.Cells(Me.Range("query_errors").row, 16384).End(xlToLeft).Column + 1).Value = "Error " & Err.Number & ": " & Err.Description
    Resume Next
End Sub
