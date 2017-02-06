' Lawson Journal Entry Tool
' Copyright (C) 2016-2017 Joe Carey
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
Private Type ProfileAttributes
    lawsonusername As String ' dddddd\uuuuuuuu
    id As String ' uuuuuuuu
    lawsonuserlogin As String ' NT00000nnn
End Type
Private Type DebitCredit
    debit As Currency
    credit As Currency
End Type
Public Sub JournalEntryReport(ByVal Co As Integer, ByVal Sys As String, ByVal JeType As String, ByVal CtrlGrp As Long, ByVal FY As Integer, ByVal Pd As Integer)

    Dim url As String
    Dim response As String
    Dim row As Long  ' Row # for report worksheet output
    Dim line As Long ' Row # for GL Detail output
    Dim detail_lines_count As Long
    Dim last_page_break As Integer
    Dim je_operator_name As String
    Dim yes_or_no As String ' Yes/No as string
    Dim auto_reverse_period As String ' Auto reverse period as string
    Dim je_type As String ' JE Type from JE Header
    Dim control_group As String ' JE Control Group from JE Header
    Dim je_sequence As String ' JE Sequence from JE Header
    Dim line_amount As DebitCredit
    Dim lines_total As DebitCredit
    Dim reversing_amount As DebitCredit
    Dim base_amount As DebitCredit
    Dim unit_amount As DebitCredit
    Dim prof As ProfileAttributes

    If Not CheckUserAttributes() Then Login
    prof = ProfileUser()

    url = "PROD=" & g_sProductLine & "&FILE=GLCONTROL&INDEX=GLCSET1"
    url = url & "&KEY=" & FilterForWeb(Co & "=" & FY & "=" & Pd & "=" & Sys & "=" & JeType & "=" & CtrlGrp)
    url = url & "&FIELD=" & FilterForWeb("COMPANY;COMPANY.NAME;COMPANY.CURRENCY-CODE;FISCAL-YEAR;ACCT-PERIOD;SYSTEM;JE-TYPE;CONTROL-GROUP;JE-SEQUENCE;DESCRIPTION;STATUS,xlt;HOLD-CODE;HOLD-REM-OPER;OPERATOR;POSTING-DATE;DATE;AUTO-REV;AUTO-REV-PD;REFERENCE;DOCUMENT-NBR;JRNL-BOOK-NBR;NBR-LINES")
    url = url & "&OUT=XML&NEXT=FALSE&MAX=1&keyUsage=PARAM"

    response = SendURL(url, "D")
    SetXMLObject ' Load IE page document into XML document object
    If Not g_oDom.LoadXML(response) Then Exit Sub
    If g_oDom.DocumentElement.SelectSingleNode("/DME") Is Nothing Then ' do we have a /DME xml document?
        Me.Range("B5").Value = GetNodeText("/ERROR/MSG", 1) ' Error message from GLCONTROL query
        Exit Sub
    End If

' Header query columns:
'  1:COMPANY       2:COMPANY.NAME  3:COMPANY.CURRENCY-CODE  4:FISCAL-YEAR     5:ACCT-PERIOD  6:SYSTEM         7:JE-TYPE  8:CONTROL-GROUP  9:JE-SEQUENCE
' 10:DESCRIPTION  11:STATUS,xlt   12:HOLD-CODE             13:HOLD-REM-OPER  14:OPERATOR    15:POSTING-DATE  16:DATE    17:AUTO-REV      18:AUTO-REV-PD
' 19:REFERENCE    20:DOCUMENT-NBR 21:JRNL-BOOK-NBR         22:NBR-LINES
    Dim rArray() As String
    Dim qArray() As String
    DmeToArray dme_array:=qArray

    detail_lines_count = Int("0" & Trim(qArray(0, 21))) ' Number of lines as reported by GLCONTROL query
    ReDim rArray(detail_lines_count * 3 + 16)
    row = 11    ' Header goes through row 11, rows 12-end vary
    DeleteRows below_range:=Me.Cells(row + 1, 1) ' delete any previous outputs and reset UsedRange

    je_type = qArray(0, 6)
    control_group = qArray(0, 7)
    je_sequence = qArray(0, 8)
    je_operator_name = qArray(0, 13)
' ##### TODO : Get human readable name for users other than the current logged in user.
    If je_operator_name = prof.lawsonuserlogin Then
        je_operator_name = prof.lawsonusername
    End If
    If qArray(0, 16) = "Y" Then yes_or_no = "Yes" Else yes_or_no = "No "
    auto_reverse_period = qArray(0, 17)
    rArray(0) = "GL240 Date " & format(Now(), "MM/DD/YY") & Space(26) & "Company " & Right(Space(4) & qArray(0, 0), 4) & _
        " - " & Left(qArray(0, 1) & Space(32), 32) & Left(qArray(0, 2) & Space(27), 27)
    rArray(1) = "      Time " & format(Now(), "HH:MM") & Space(29) & "Journal Edit Listing"
    rArray(2) = Space(45) & "For Fiscal Year " & qArray(0, 3) & " - Periods " & format(qArray(0, 4), "00") & _
        " - " & format(qArray(0, 4), "00")
    rArray(3) = ""
    rArray(4) = " Journal            " & qArray(0, 5) & " " & je_type & _
        Right("         " & control_group, 9) & "-" & format(je_sequence, "00") & " " & Left(qArray(0, 9) & Space(36), 36) & _
        "Fiscal Year    " & qArray(0, 3) & "          Period     " & qArray(0, 4)
    rArray(5) = "   Status           " & Left(qArray(0, 10) & Space(17), 17) & "Hold Code " & _
        Left(qArray(0, 11) & Space(10), 10) & "Hold Removal Operator " & Left(qArray(0, 12) & Space(13), 13) & _
        "Operator       " & je_operator_name
    rArray(6) = "   Posting Date     " & format(qArray(0, 14), "MM/DD/YY") & _
        "         Transaction Date    " & format(qArray(0, 15), "MM/DD/YY") & "          Reverse  " & yes_or_no & "     Reverse Pd     " & auto_reverse_period
    rArray(7) = "   Reference        " & Left(qArray(0, 18) & Space(16), 16) & " Document            " & _
        Left(qArray(0, 19) & Space(34), 34) & " Journal Book   " & Left(qArray(0, 20) & Space(9), 9)
    rArray(8) = ""
    rArray(9) = " Line   Co           Account                     Activity         Ref    SC Rvs                    Debit                    Credit"
    rArray(10) = "------ ---- --------------------------- --------------------- ---------- -- --- ------------------------- -------------------------"

    url = "PROD=" & g_sProductLine ' Start building POST data string with Product Line
    url = url & "&FILE=GLTRANS&INDEX=GLTSET1" ' ' Table GLTRANS, criteria set GLTSET1: key = co=fy=pd=ctrlgrp=sys=jetype
    url = url & "&KEY=" & FilterForWeb(Co & "=" & FY & "=" & Pd & "=" & CtrlGrp & "=" & Sys & "=" & JeType)
    url = url & "&FIELD=" & FilterForWeb("LINE-NBR;TO-COMPANY;ACCT-UNIT;ACCOUNT;SUB-ACCOUNT;ACTIVITY;ACCT-CATEGORY;REFERENCE;SOURCE-CODE;AUTO-REV;TRAN-AMOUNT;CHART-DETAIL.ACCOUNT-DESC;DESCRIPTION;BASE-AMOUNT;UNITS-AMOUNT")
    url = url & "&OUT=XML&NEXT=FALSE&MAX=10000&keyUsage=PARAM"

    response = SendURL(url, "D")
    If Not g_oDom.LoadXML(response) Then Exit Sub
    If g_oDom.DocumentElement.SelectSingleNode("/DME") Is Nothing Then ' do we have a /DME xml document?
        Me.Range("B6").Value = GetNodeText("/ERROR/MSG", 1) ' Error message from GLCONTROL query
        Exit Sub
    End If

    base_amount = dcSet(0, 0)
    reversing_amount = dcSet(0, 0)
    lines_total = dcSet(0, 0)
    unit_amount = dcSet(0, 0)
    Call DmeToArray(qArray)

' Detail lines query columns:
'  1:LINE-NBR      2:TO-COMPANY   3:ACCT-UNIT    4:ACCOUNT       5:SUB-ACCOUNT   6:ACTIVITY     7:ACCT-CATEGORY  8:REFERENCE  9:SOURCE-CODE  10:AUTO-REV
' 11:TRAN-AMOUNT  12:CHART-DETAIL.ACCOUNT-DESC  13:DESCRIPTION  14:BASE-AMOUNT  15:UNITS-AMOUNT
    For line = 0 To detail_lines_count - 1
        line_amount = CurrencyToDebitCredit(qArray(line, 10))
        lines_total = dcAdd(lines_total, line_amount) ' Add line amounts for total
        base_amount = dcAdd(base_amount, CurrencyToDebitCredit(qArray(line, 13))) ' Add Base Amounts for total
        unit_amount = dcAdd(unit_amount, CurrencyToDebitCredit(qArray(line, 14))) ' Add Unit Amounts for total

        If qArray(line, 9) = "Y" Then  ' Set "Yes", "No" or blank string for Auto-Reverse line and add to total
            yes_or_no = "Yes"
            reversing_amount = dcAdd(reversing_amount, line_amount)
        ElseIf qArray(line, 9) = "N" Then
            yes_or_no = "No "
        Else
            yes_or_no = "   "
        End If

        If qArray(line, 0) <> "" Then
            rArray(row) = Right("      " & qArray(line, 0), 6) & " " & Right("    " & qArray(line, 1), 4) & " " & _
                Left(qArray(line, 2) & Space(17), 17) & format(qArray(line, 3), "00000") & "-" & format(qArray(line, 4), "0000") & _
                " " & Left(qArray(line, 5) & Space(16), 16) & Left(qArray(line, 6) & Space(5), 5) & " " & Left(qArray(line, 7) & _
                Space(10), 10) & " " & Left(qArray(line, 8) & "  ", 2) & " " & yes_or_no & Right(Space(26) & format(line_amount.debit, "#,##0.00;;\ "), 26) & _
                Right(Space(26) & format(line_amount.credit, "#,##0.00;;\ "), 26)
            rArray(row + 1) = Left(qArray(line, 11) & Space(27), 27) & " " & qArray(line, 12)
            rArray(row + 2) = ""
            row = row + 3
        End If
    Next line
    rArray(row) = "*** Totals For Journal entry " & je_type & "-" & Right("       " & control_group, 8) & "-" & _
        format(je_sequence, "00") & "                              Debits                   Credits                Difference"
    rArray(row + 1) = Space(37) & "Base . . . . . . ." & Right(Space(24) & format(base_amount.debit, "#,##0.00"), 24) & _
        Right(Space(26) & format(base_amount.credit, "#,##0.00"), 26) & Right(Space(26) & format(dcNet(base_amount), "#,##0.00"), 26)
    rArray(row + 2) = Space(37) & "Reverse  . . . . ." & Right(Space(24) & format(reversing_amount.debit, "#,##0.00"), 24) & _
        Right(Space(26) & format(reversing_amount.credit, "#,##0.00"), 26) & Right(Space(26) & format(dcNet(reversing_amount), "#,##0.00"), 26)
    rArray(row + 3) = Space(37) & "Entered  . . . . ." & Right(Space(24) & format(lines_total.debit, "#,##0.00"), 24) & _
        Right(Space(26) & format(lines_total.credit, "#,##0.00"), 26) & Right(Space(26) & format(dcNet(lines_total), "#,##0.00"), 26)
    rArray(row + 4) = Space(37) & "Unit . . . . . . ." & Right(Space(24) & format(unit_amount.debit, "#,##0.00"), 24) & _
        Right(Space(26) & format(unit_amount.credit, "#,##0.00"), 26) & Right(Space(26) & format(dcNet(unit_amount), "#,##0.00"), 26)

    ' Move array data onto worksheet
    line = UBound(rArray)
    Me.Range("A1").Resize(line).Value = Application.Transpose(rArray)

' Output relies on worksheet print format to print correctly:
'   8.5 x 11 landscape; fit to 1 page wide; print area: $A:$A; rows to repeat at top: $1:$11
'   margins: top: 0.55 bottom: 0.70 left: 0.25 right: 0.25 header: 0.25 footer: 0.00
'   header left: "Journal Edit Listing" Arial 16 right: "Page &[Page] of &[Pages] " Courier New 9
' #### TODO : Set header content
    With Me.PageSetup
        .Orientation = xlLandscape
        .FitToPagesWide = 1
        .PrintArea = "$A:$A"
        .PrintTitleRows = "$1:$11"
        .TopMargin = Application.InchesToPoints(0.55)
        .BottomMargin = Application.InchesToPoints(0.7)
        .LeftMargin = Application.InchesToPoints(0.25)
        .RightMargin = Application.InchesToPoints(0.25)
        .HeaderMargin = Application.InchesToPoints(0.25)
        .FooterMargin = 0
    End With

    ' Ensure totals block (last 5 lines) is on one page
    row = line - 4 ' Start row of totals block
    Me.ResetAllPageBreaks
    last_page_break = Me.HPageBreaks.Count
    If last_page_break > 0 Then
        If Me.HPageBreaks.Item(last_page_break).Location.row > row Then
            Me.HPageBreaks.Add (Me.Rows(row))
        End If
    End If

End Sub
Private Function CurrencyToDebitCredit(ByVal input_currency As String) As DebitCredit
    Dim input_is_negative As Boolean
    input_is_negative = input_currency Like "*-"
    input_currency = Replace(input_currency, "-", "")
    input_currency = Replace(input_currency, " ", "")
    CurrencyToDebitCredit.debit = 0
    CurrencyToDebitCredit.credit = 0
    If input_is_negative Then
        CurrencyToDebitCredit.credit = Val(input_currency)
    Else
        CurrencyToDebitCredit.debit = Val(input_currency)
    End If
End Function
Private Function dcNet(ByRef a As DebitCredit) As Currency
    dcNet = a.debit - a.credit
End Function
Private Function dcSet(ByVal a As Currency, ByVal b As Currency) As DebitCredit
    dcSet.debit = a
    dcSet.credit = b
End Function
Private Function dcAdd(ByRef a As DebitCredit, ByRef b As DebitCredit) As DebitCredit
    dcAdd.debit = a.debit + b.debit
    dcAdd.credit = a.credit + b.credit
End Function
Private Function ProfileUser() As ProfileAttributes
    Dim url As String
    Dim response As String
    Dim xml_node As MSXML2.IXMLDOMNode
    Dim xml_node_list As MSXML2.IXMLDOMNodeList

    url = g_sServer & "/servlet/Profile?section=attributes"
    response = SendURL(url, "X")

    If Not g_oDom.LoadXML(response) Then
        ProfileUser.lawsonusername = ""
        Exit Function
    End If
    If g_oDom.DocumentElement.SelectSingleNode("/PROFILE") Is Nothing Then
        ProfileUser.lawsonusername = ""
        Exit Function
    End If

    Set xml_node_list = g_oDom.DocumentElement.SelectNodes("//ATTR")
    For Each xml_node In xml_node_list
        Select Case xml_node.Attributes.getNamedItem("name").Text
            Case "lawsonusername"
                ProfileUser.lawsonusername = xml_node.Attributes.getNamedItem("value").Text
            Case "id"
                ProfileUser.id = xml_node.Attributes.getNamedItem("value").Text
            Case "lawsonuserlogin"
                ProfileUser.lawsonuserlogin = xml_node.Attributes.getNamedItem("value").Text
        End Select
    Next xml_node

End Function
Private Sub DeleteRows(ByVal below_range As Range)
On Error GoTo error_handler
    Dim original_active_sheet As Worksheet
    Set original_active_sheet = ActiveSheet

    Dim target_sheet As Worksheet
    Set target_sheet = below_range.Worksheet

    Dim last_used_row As Long
    last_used_row = target_sheet.UsedRange.Rows.Count ' last row in UsedRange

    Dim below_range_row As Long
    below_range_row = below_range.row

    target_sheet.Activate ' activate destination worksheet to faciliate resetting UsedRange
    If last_used_row > below_range_row Then ' if our destination row is after UsedRange, we'd end up deleting the last row of UsedRange (e.g., rows("10:2").delete)
        below_range.Worksheet.Rows(below_range_row & ":" & last_used_row).Delete
    End If
    Application.ActiveSheet.UsedRange ' Reset the Used Range to control file size
    original_active_sheet.Activate
    Exit Sub
error_handler:
    ' if we hit an error deleting rows, just continue with whatever doesn't error
    Debug.Print "Function DeleteRows error " & Err.Number & " (" & Err.description & ")"
    Resume Next
End Sub
Private Sub DmeToArray(ByRef dme_array() As String)
    Dim records_node_list As MSXML2.IXMLDOMNodeList
    Dim record_node As MSXML2.IXMLDOMNode
    Dim records_count As Integer
    Dim record_index As Integer
    Set records_node_list = g_oDom.SelectNodes("/DME/RECORDS/RECORD/COLS")
    records_count = records_node_list.Length

    Dim columns_node_list As MSXML2.IXMLDOMNodeList
    Dim column_node As MSXML2.IXMLDOMNode
    Dim columns_count As Integer
    Dim column_index As Integer
    columns_count = g_oDom.SelectNodes("/DME/COLUMNS/COLUMN").Length

    Dim item_value As String

    ReDim dme_array(records_count, columns_count)
    record_index = 0
    For Each record_node In records_node_list
        Set columns_node_list = record_node.SelectNodes("COL")
        column_index = 0
        For Each column_node In columns_node_list
            item_value = column_node.FirstChild.Text
            If Len(item_value) > 0 Then
                dme_array(record_index, column_index) = item_value
            End If
            column_index = column_index + 1
        Next column_node
        record_index = record_index + 1
    Next record_node
End Sub
Private Function GetNodeText(ByVal base_node_xpath As String, ByVal index As Integer, Optional ByVal child_node_xpath = "/") As String
    Dim xml_node As MSXML2.IXMLDOMNode
    Dim xml_node_list As MSXML2.IXMLDOMNodeList
    Dim xml_child_path As String
    If child_node_xpath = "/" Then
        xml_child_path = ""
    Else
        xml_child_path = "/" & child_node_xpath
    End If
    Set xml_node_list = g_oDom.DocumentElement.SelectNodes(base_node_xpath)
    If xml_node_list.Length > 0 And xml_node_list.Length >= index Then
        Set xml_node = g_oDom.DocumentElement.SelectSingleNode(base_node_xpath & "[" & index & "]" & xml_child_path)
        GetNodeText = xml_node.FirstChild.Text
    Else
        GetNodeText = ""
    End If
End Function
