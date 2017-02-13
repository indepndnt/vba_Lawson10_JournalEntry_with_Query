' Lawson Journal Entry Tool
' Copyright (C) 2017 Joe Carey
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
Private Type JournalHeader
    exists As Boolean
    function_code As String
    company As Integer
    fiscal_year As Integer
    acct_period As Integer
    system As String
    je_type As String
    auto_rev As String
    control_group As Integer
    je_sequence As Integer
    description As String
    post_date As Date
    source_code As String
    auto_rev_pd As Integer
    tran_date As Date
    reference As String
    document_nbr As String
    function_code_cell As Range
    control_group_cell As Range
    je_sequence_cell As Range
    description_cell As Range
    response_cell As Range
End Type
Private headers() As JournalHeader
Private Sub JournalsUpload_Click()
' On Error GoTo error_handler
    If Not CheckUserAttributes() Then Login
    UploadJournalHeader
    UploadJournalDetails
    FixObjects
    Exit Sub
error_handler:
    MsgBox ("Upload Error" & vbCrLf & Err.Number & ":" & Err.description)
End Sub
Private Sub FixObjects()
    ' Help from http://stackoverflow.com/questions/19385803/how-to-stop-activex-objects-automatically-changing-size-in-office
    ' JournalsUpload
    Me.JournalsUpload.Left = 408.75
    Me.JournalsUpload.Width = 60
    Me.JournalsUpload.Height = 21.75
    Me.JournalsUpload.Top = 9.75
    Me.Shapes("JournalsUpload").ScaleHeight 1.25, msoFalse, msoScaleFromTopLeft
    Me.Shapes("JournalsUpload").ScaleHeight 0.8, msoFalse, msoScaleFromTopLeft

    ' auto_clear_numbers
    Me.auto_clear_numbers.Left = 397.5
    Me.auto_clear_numbers.Width = 150
    Me.auto_clear_numbers.Height = 18.75
    Me.auto_clear_numbers.Top = 33
    Me.Shapes("auto_clear_numbers").ScaleHeight 1.25, msoFalse, msoScaleFromTopLeft
    Me.Shapes("auto_clear_numbers").ScaleHeight 0.8, msoFalse, msoScaleFromTopLeft
End Sub
Private Sub ReadJournalHeaders()
    Dim index As Integer
    Dim row As Integer
    Dim number_of_rows As Integer
    Dim cell_range As Range

    Dim header_column_company As Integer
    Dim header_column_fiscal_year As Integer
    Dim header_column_acct_period As Integer
    Dim header_column_system As Integer
    Dim header_column_je_type As Integer
    Dim header_column_auto_rev As Integer
    Dim header_column_control_group As Integer
    Dim header_column_je_sequence As Integer
    Dim header_column_description As Integer
    Dim header_column_posting_date As Integer
    Dim header_column_source_code As Integer
    Dim header_column_auto_rev_pd As Integer
    Dim header_column_date As Integer
    Dim header_column_reference As Integer
    Dim header_column_document_nbr As Integer
    Dim header_column_response As Integer

    ' Initialize header columns
    header_column_company = 0
    header_column_fiscal_year = 0
    header_column_acct_period = 0
    header_column_system = 0
    header_column_je_type = 0
    header_column_auto_rev = 0
    header_column_control_group = 0
    header_column_je_sequence = 0
    header_column_description = 0
    header_column_posting_date = 0
    header_column_source_code = 0
    header_column_auto_rev_pd = 0
    header_column_date = 0
    header_column_reference = 0
    header_column_document_nbr = 0
    header_column_response = 0
    row = Me.Range("Start_Header_Section").row + 1
    number_of_rows = Me.Range("Start_JE_Detail_Section").row - row - 1

    ' size headers array to number of header rows in worksheet
    ReDim headers(number_of_rows) As JournalHeader

    For Each cell_range In Range(Me.Cells(row, 1), Me.Cells(row, 16384).End(xlToLeft))
        Select Case cell_range.Value2
            Case "COMPANY"
                header_column_company = cell_range.Column
            Case "FISCAL-YEAR"
                header_column_fiscal_year = cell_range.Column
            Case "ACCT-PERIOD"
                header_column_acct_period = cell_range.Column
            Case "SYSTEM"
                header_column_system = cell_range.Column
            Case "JE-TYPE"
                header_column_je_type = cell_range.Column
            Case "AUTO-REV"
                header_column_auto_rev = cell_range.Column
            Case "CONTROL-GROUP"
                header_column_control_group = cell_range.Column
            Case "JE-SEQUENCE"
                header_column_je_sequence = cell_range.Column
            Case "DESCRIPTION"
                header_column_description = cell_range.Column
            Case "POSTING-DATE"
                header_column_posting_date = cell_range.Column
            Case "SOURCE-CODE"
                header_column_source_code = cell_range.Column
            Case "AUTO-REV-PD"
                header_column_auto_rev_pd = cell_range.Column
            Case "DATE"
                header_column_date = cell_range.Column
            Case "REFERENCE"
                header_column_reference = cell_range.Column
            Case "DOCUMENT-NBR"
                header_column_document_nbr = cell_range.Column
            Case "Response"
                header_column_response = cell_range.Column
        End Select
    Next cell_range
    If header_column_control_group = 0 Or header_column_response = 0 Or header_column_company = 0 Or header_column_je_type = 0 Or _
        (header_column_posting_date = 0 And (header_column_fiscal_year = 0 Or header_column_acct_period = 0)) Or _
        header_column_system = 0 Or header_column_description = 0 Then
        MsgBox "Missing required header column!  (COMPANY, JE-TYPE, POSTING-DATE (or FISCAL-YEAR and ACCT-PERIOD), SYSTEM, DESCRIPTION, CONTROL-GROUP, Response)"
        Exit Sub
    End If

    ' Loop through header rows
    For index = 1 To number_of_rows
        row = Me.Range("Start_Header_Section").row + index + 1

        headers(index).exists = True
        Set headers(index).response_cell = Me.Cells(row, header_column_response)

        Set headers(index).function_code_cell = Me.Cells(row, 1)
        headers(index).function_code = headers(index).function_code_cell.Value2

        headers(index).company = 0
        If header_column_company > 0 Then
            If Me.Cells(row, header_column_company).Value2 <> "" And IsNumeric(Me.Cells(row, header_column_company).Value2) Then
                headers(index).company = Me.Cells(row, header_column_company).Value2
            End If
        End If

        If Me.Cells(row, header_column_fiscal_year).Value2 <> "" And IsNumeric(Me.Cells(row, header_column_fiscal_year).Value2) Then
            headers(index).fiscal_year = Me.Cells(row, header_column_fiscal_year).Value2
        Else
            headers(index).fiscal_year = 0
        End If

        headers(index).acct_period = 0
        If header_column_acct_period > 0 Then
            If Me.Cells(row, header_column_acct_period).Value2 <> "" And IsNumeric(Me.Cells(row, header_column_acct_period).Value2) Then
                headers(index).acct_period = Me.Cells(row, header_column_acct_period).Value2
            End If
        End If

        If Me.Cells(row, header_column_system).Value2 <> "" Then headers(index).system = Me.Cells(row, header_column_system).Value2 Else headers(index).system = "GL"
        If Me.Cells(row, header_column_je_type).Value2 <> "" Then headers(index).je_type = Me.Cells(row, header_column_je_type).Value2 Else headers(index).je_type = ""

        headers(index).auto_rev = ""
        If header_column_auto_rev > 0 Then
            If Me.Cells(row, header_column_auto_rev).Value2 <> "" Then headers(index).auto_rev = Me.Cells(row, header_column_auto_rev).Value2
        End If

        Set headers(index).control_group_cell = Me.Cells(row, header_column_control_group)
        If headers(index).control_group_cell.Value2 <> "" And IsNumeric(headers(index).control_group_cell.Value2) Then
            headers(index).control_group = headers(index).control_group_cell.Value2
        Else
            headers(index).control_group = 0
        End If

        headers(index).je_sequence = 0
        If header_column_je_sequence > 0 Then
            Set headers(index).je_sequence_cell = Me.Cells(row, header_column_je_sequence)
            If headers(index).je_sequence_cell.Value2 <> "" And IsNumeric(headers(index).je_sequence_cell.Value2) Then
                headers(index).je_sequence = headers(index).je_sequence_cell.Value2
            End If
        End If

        Set headers(index).description_cell = Me.Cells(row, header_column_description)
        If headers(index).description_cell.Value2 <> "" Then headers(index).description = headers(index).description_cell.Value2 Else headers(index).description = ""

        headers(index).post_date = 0
        If header_column_posting_date > 0 Then
            If IsDate(Me.Cells(row, header_column_posting_date).Value2) Then
                headers(index).post_date = Me.Cells(row, header_column_posting_date).Value2
            End If
        End If

        headers(index).source_code = ""
        If header_column_source_code > 0 Then
            If Me.Cells(row, header_column_source_code).Value2 <> "" Then headers(index).source_code = Me.Cells(row, header_column_source_code).Value2
        End If

        headers(index).auto_rev_pd = 0
        If header_column_auto_rev_pd > 0 Then
            If Me.Cells(row, header_column_auto_rev_pd).Value2 <> "" And IsNumeric(Me.Cells(row, header_column_auto_rev_pd).Value2) Then
                headers(index).auto_rev_pd = Me.Cells(row, header_column_auto_rev_pd).Value2
            End If
        End If

        headers(index).tran_date = 0
        If header_column_date > 0 Then
            If IsDate(Me.Cells(row, header_column_date).Value2) Then
                headers(index).tran_date = Me.Cells(row, header_column_date).Value2
            End If
        End If

        headers(index).reference = ""
        If header_column_reference > 0 Then
            If Me.Cells(row, header_column_reference).Value2 <> "" Then
                headers(index).reference = Me.Cells(row, header_column_reference).Value2
            End If
        End If

        headers(index).document_nbr = ""
        If header_column_document_nbr > 0 Then
            If Me.Cells(row, header_column_document_nbr).Value2 <> "" Then
                headers(index).document_nbr = Me.Cells(row, header_column_document_nbr).Value2
            End If
        End If

        If headers(index).post_date = 0 Then
            If headers(index).fiscal_year <> 0 And headers(index).acct_period <> 0 Then
                headers(index).post_date = DateSerial(headers(index).fiscal_year, headers(index).acct_period, 1)
            ElseIf headers(index).function_code <> "" Then
                headers(index).response_cell.Value2 = "Must have either POSTING-DATE or FISCAL-YEAR _and_ ACCT-PERIOD."
            End If
        ElseIf headers(index).fiscal_year = 0 Or headers(index).acct_period = 0 Then
            headers(index).fiscal_year = Year(headers(index).post_date)
            headers(index).acct_period = Month(headers(index).post_date)
        End If
    Next index
End Sub
Private Sub UploadJournalHeader()
    Const search_string As String = "//text()"      ' XPath for all text elements in the document
    Dim response_element As MSXML2.IXMLDOMNode
    Dim request_parameters As String         ' String for building POST data
    Dim request_response As String
    Dim status As Integer   ' Status number return value
    Dim message_code As Integer     ' Message number return value
    Dim idx As Integer

    ReadJournalHeaders

    For idx = 1 To UBound(headers())
        status = 0
        ' build request string
        request_parameters = "_PDL=" & g_sProductLine
        Select Case headers(idx).function_code
            Case "A"        ' Add - ensure there's no JE #
                If headers(idx).control_group  = 0 Or Me.auto_clear_numbers Then
                    headers(idx).control_group_cell.Value2 = ""
                    headers(idx).control_group = 0
                    request_parameters = request_parameters & "&_TKN=GL40.2&_EVT=ADD&_RTN=DATA&_TDS=IGNORE&FC=Add"
                    headers(idx).response_cell.Value2 = ""  ' Clear the system response cell for the Header upload row on the spreadsheet
                Else
                    headers(idx).response_cell.Value2 = "To add new, JE# must be blank."
                    status = 999
                End If
            Case "C"        ' Change - ensure there is a JE #
                If headers(idx).control_group > 0 Then
                    request_parameters = request_parameters & "&_TKN=GL40.2&_EVT=CHG&_RTN=DATA&_TDS=IGNORE&FC=Change"
                    headers(idx).response_cell.Value2 = ""  ' Clear the system response cell for the Header upload row on the spreadsheet
                Else
                    headers(idx).response_cell.Value2 = "To change JE header, must specify JE#."
                    status = 999
                End If
            Case "D"        ' Delete - ensure there is a JE #
                If headers(idx).control_group > 0 Then
                    ' ' deleting JE header requires hidden key (HK) value with Company, FY, Period, System, JE Type, JE #, and JE seq # formatted as 24 characters: ccccyyyymmsstccccccccqq
                    request_parameters = request_parameters & "&_TKN=GL40.2&_EVT=CHG&_RTN=DATA&_TDS=IGNORE&FC=Delete&HK=" & _
                        format(headers(idx).company, "0000") & format(headers(idx).post_date, "yyyymm") & headers(idx).system & _
                        headers(idx).je_type & format(headers(idx).control_group, "00000000") & format(headers(idx).je_sequence, "00")
                    headers(idx).response_cell.Value2 = ""  ' Clear the system response cell for the Header upload row on the spreadsheet
                Else
                    headers(idx).response_cell.Value2 = "To delete JE header, must specify JE#."
                    status = 999
                End If
            Case ""         ' blank - skip header, continue to detail (add or change)
                status = 999
            Case Else       ' not 'A' or 'C' or blank??
                headers(idx).response_cell.Value2 = "Unknown function code - 'A', 'C' or 'D' only, blank to skip."
                status = 999
        End Select

        With headers(idx)
            request_parameters = request_parameters & "&_f17=" & .company & "&_f20=" & .fiscal_year & _
                "&_f21=" & .acct_period & "&_f22=" & .system & "&_f24=" & .je_type
            If .control_group <> 0 Then request_parameters = request_parameters & "&_f25=" & .control_group ' will error if FC=A
            If .je_sequence <> 0 Then request_parameters = request_parameters & "&_f26=" & .je_sequence
            request_parameters = request_parameters & "&_f27=" & FilterForWeb(Left(.description, 30)) ' (reqired) 30 characters max field length
            If .source_code <> "" Then request_parameters = request_parameters & "&_f30=" & .source_code ' defaults to "JE"
            If .reference <> "" Then request_parameters = request_parameters & "&_f34=" & FilterForWeb(.reference)
            If .auto_rev <> "" Then request_parameters = request_parameters & "&_f37=" & .auto_rev ' defaults to "N"
            If .auto_rev_pd <> 0 Then request_parameters = request_parameters & "&_f38=" & .auto_rev_pd ' defaults to 0-next pd
            If .document_nbr <> "" Then request_parameters = request_parameters & "&_f42=" & FilterForWeb(.document_nbr)
            request_parameters = request_parameters & "&_f48=" & format(.post_date, "yyyymmdd") ' (required)
            If .tran_date <> 0 Then request_parameters = request_parameters & "&_f49=" & format(.tran_date, "yyyymmdd")
            request_parameters = request_parameters & "&_OUT=XML" ' "&_EOT=TRUE"
        End With

        If status = 0 And headers(idx).company <> 0 And headers(idx).post_date <> 0 And headers(idx).system <> "" And headers(idx).je_type <> "" And headers(idx).description <> "" Then
            SetXMLObject ' Load page document into XML document object
            request_response = SendURL(request_parameters, "T")
            If Not g_oDom.LoadXML(request_response) Then
                If headers(idx).function_code = "A" Then
                    headers(idx).function_code_cell.Value2 = "C"
                    headers(idx).response_cell.Value2 = "Loading error - check if JE header exists before adding again."
                Else
                    headers(idx).response_cell.Value2 = "Loading error - check JE report to confirm change."
                End If
            Else
                For Each response_element In g_oDom.SelectNodes(search_string) ' Decide what to do with each text element based on the parent node's name
                    Select Case response_element.ParentNode.nodeName
                        Case "Message"      ' This is the node name that returns the upload result
                            headers(idx).response_cell.Value2 = headers(idx).response_cell.Value2 & response_element.Text
                        Case "FldNbr"       ' This is the field number related to the response message
                            headers(idx).response_cell.Value2 = headers(idx).response_cell.Value2 & "(" & response_element.Text & ")"
                        Case "MsgNbr"       ' This return value refers to the message response - message number "000" is success
                            message_code = Val(response_element.Text)
                        Case "StatusNbr"    ' This is the status of the header upload.
                            status = Val(response_element.Text)
                        Case "_f25"         ' This is the Control Group (JE #)
                            headers(idx).control_group = Val(response_element.Text)
                            headers(idx).control_group_cell.Value2 = response_element.Text
                        Case "_f26"         ' This is the JE Sequence Num
                            headers(idx).je_sequence = Val(response_element.Text)
                            headers(idx).je_sequence_cell.Value2 = response_element.Text
                    End Select
                Next response_element
            End If

            If status = 1 And message_code = 0 Then ' If add/change/delete successful, delete the Function Code
                If headers(idx).function_code = "D" Then headers(idx).control_group_cell.Value2 = "deleted (" & headers(idx).control_group & ")" ' mark if deleted
                headers(idx).function_code_cell.Value2 = ""
                ' Add hyperlink to report
                If Not ReportSheet Is ActiveSheet And headers(idx).description <> "" Then
                    headers(idx).description_cell.Hyperlinks.Add Anchor:=headers(idx).description_cell, Address:="", SubAddress:="Report!$A$12", _
                        ScreenTip:="Journal Entry Report [" & Join(Array(headers(idx).company, headers(idx).system, headers(idx).je_type, _
                        headers(idx).control_group, headers(idx).fiscal_year, headers(idx).acct_period), ";") & "]"
                End If
                Me.Calculate ' calculate worksheet - in case detail line formulas are picking up the header.control_group
            End If
        End If
    Next idx
End Sub
Private Function GetJournalHeader(ByVal company As Integer, ByVal fiscal_year As Integer, ByVal acct_period As Integer, ByVal system As String, ByVal je_type As String, ByVal control_group As Integer) As JournalHeader
    Dim index As Integer
    GetJournalHeader.exists = False
    For index = 1 To UBound(headers)
        If headers(index).company = company _
            And headers(index).fiscal_year = fiscal_year _
            And headers(index).acct_period = acct_period _
            And headers(index).system = system _
            And headers(index).je_type = je_type _
            And headers(index).control_group = control_group Then GetJournalHeader = headers(index)
    Next index
End Function
Private Sub UploadJournalDetails()
    Dim response_element As MSXML2.IXMLDOMNode
    Dim request_parameters As String            ' String for building POST data
    Dim request_response As String
    Dim status As Integer                       ' Status number return value
    Dim message_code As Integer                 ' Message number return value
    Const search_string As String = "//text()"  ' XPath for all text elements in the document

    Dim row As Long
    Dim cell_range As Range
    Dim detail_range As Range
    Dim je As JournalHeader

    Dim detail_column_company As Integer
    Dim detail_column_fiscal_year As Integer
    Dim detail_column_acct_period As Integer
    Dim detail_column_system As Integer
    Dim detail_column_je_type As Integer
    Dim detail_column_auto_rev As Integer
    Dim detail_column_control_group As Integer
    Dim detail_column_to_company As Integer
    Dim detail_column_description As Integer
    Dim detail_column_line_nbr As Integer
    Dim detail_column_acct_unit As Integer
    Dim detail_column_account As Integer
    Dim detail_column_sub_account As Integer
    Dim detail_column_reference As Integer
    Dim detail_column_activity As Integer
    Dim detail_column_acct_category As Integer
    Dim detail_column_tran_amount As Integer
    Dim detail_column_response As Integer

    ' Initialize detail columns
    detail_column_company = 0
    detail_column_fiscal_year = 0
    detail_column_acct_period = 0
    detail_column_system = 0
    detail_column_je_type = 0
    detail_column_auto_rev = 0
    detail_column_control_group = 0
    detail_column_to_company = 0
    detail_column_description = 0
    detail_column_line_nbr = 0
    detail_column_acct_unit = 0
    detail_column_account = 0
    detail_column_sub_account = 0
    detail_column_reference = 0
    detail_column_activity = 0
    detail_column_acct_category = 0
    detail_column_tran_amount = 0
    detail_column_response = 0
    row = Me.Range("Start_JE_Detail_Section").row + 1
    Set detail_range = Range(Me.Cells(row, 1), Me.Cells(row, 16384).End(xlToLeft))
    For Each cell_range In detail_range
        Select Case cell_range.Value2
            Case "COMPANY"
                detail_column_company = cell_range.Column
            Case "FISCAL-YEAR"
                detail_column_fiscal_year = cell_range.Column
            Case "ACCT-PERIOD"
                detail_column_acct_period = cell_range.Column
            Case "SYSTEM"
                detail_column_system = cell_range.Column
            Case "JE-TYPE"
                detail_column_je_type = cell_range.Column
            Case "CONTROL-GROUP"
                detail_column_control_group = cell_range.Column
            Case "TO-COMPANY"
                detail_column_to_company = cell_range.Column
            Case "LINE-NBR"
                detail_column_line_nbr = cell_range.Column
            Case "ACCT-UNIT"
                detail_column_acct_unit = cell_range.Column
            Case "ACCOUNT"
                detail_column_account = cell_range.Column
            Case "SUB-ACCOUNT"
                detail_column_sub_account = cell_range.Column
            Case "ACTIVITY"
                detail_column_activity = cell_range.Column
            Case "ACCT-CATEGORY"
                detail_column_acct_category = cell_range.Column
            Case "AUTO-REV"
                detail_column_auto_rev = cell_range.Column
            Case "TRAN-AMOUNT"
                detail_column_tran_amount = cell_range.Column
            Case "DESCRIPTION"
                detail_column_description = cell_range.Column
            Case "REFERENCE"
                detail_column_reference = cell_range.Column
            Case "Response"
                detail_column_response = cell_range.Column
        End Select
    Next cell_range
    If detail_column_control_group = 0 Or detail_column_response = 0 Or detail_column_line_nbr = 0 Or detail_column_acct_unit = 0 Or detail_column_account = 0 Or detail_column_tran_amount = 0 Then
        MsgBox "Missing required detail column!  (ACCT-UNIT, ACCOUNT, TRAN-AMOUNT, LINE-NBR, CONTROL-GROUP, Response)"
        Exit Sub
    End If
    If detail_column_company = 0 Or detail_column_fiscal_year = 0 Or detail_column_acct_period = 0 Or detail_column_system = 0 Or detail_column_je_type = 0 Then
        MsgBox "Missing required detail column for header information!  (COMPANY, FISCAL-YEAR, ACCT-PERIOD, SYSTEM, JE-TYPE, CONTROL-GROUP)"
        Exit Sub
    End If

    ' loop through detail rows
    For row = row + 1 To Me.UsedRange.Rows.Count
        status = 0 ' Status 0 = continue on current row
        ' build request string
        request_parameters = "_PDL=" & g_sProductLine
        Select Case Me.Cells(row, 1).Value2 ' Decide how to treat line based on Function Code
            Case "A"        ' Add - ensure there's no Line #
                If Me.Cells(row, detail_column_line_nbr).Value2 = "" Or Me.auto_clear_numbers Then
                    Me.Cells(row, detail_column_line_nbr).Value2 = ""
                    request_parameters = request_parameters & "&_TKN=GL40.1&_EVT=CHG&_RTN=DATA&_TDS=IGNORE&FC=Change"
                    Me.Cells(row, detail_column_response).Value2 = "" ' Clear the system response cell for the upload row
                Else
                    Me.Cells(row, detail_column_response).Value2 = "To add new, Line # must be blank."
                    status = 1 ' Status 1 = exit loop
                End If
            Case "C", "D"  ' Change or Delete - ensure there is a Line #
                If Me.Cells(row, detail_column_line_nbr).Value2 <> "" Then
                    request_parameters = request_parameters & "&_TKN=GL40.1&_EVT=CHG&_RTN=DATA&_TDS=IGNORE&FC=Change"
                    Me.Cells(row, detail_column_response).Value2 = "" ' Clear the system response cell for the upload row
                Else
                    Me.Cells(row, detail_column_response).Value2 = "To change or delete JE line, must specify Line #."
                    status = 1 ' Status 1 = exit loop
                End If
            Case ""         ' blank - skip line
                status = 1 ' Status 1 = exit loop
            Case Else       ' not 'A' or 'C' or blank??
                Me.Cells(row, detail_column_response).Value2 = "Unknown function code - 'A', 'C' or 'D' only, blank to skip."
                status = 1 ' Status 1 = exit loop
        End Select

        If status = 0 Then ' if we haven't set status = 1 then continue with row upload
            je = GetJournalHeader(company:=Me.Cells(row, detail_column_company).Value2, _
                fiscal_year:=Me.Cells(row, detail_column_fiscal_year).Value2, _
                acct_period:=Me.Cells(row, detail_column_acct_period).Value2, _
                system:=Me.Cells(row, detail_column_system).Value2, _
                je_type:=Me.Cells(row, detail_column_je_type).Value2, _
                control_group:=Me.Cells(row, detail_column_control_group).Value2)

            request_parameters = request_parameters & "&_f39=" & je.company & "&_f44=" & je.fiscal_year & "&_f45=" & je.acct_period
            request_parameters = request_parameters & "&_f46=" & je.system & "&_f48=" & je.je_type & "&_f49=" & je.control_group
            If je.je_sequence <> 0 Then request_parameters = request_parameters & "&_f50=" & je.je_sequence ' JE Sequence # from Header, if there is one
            request_parameters = request_parameters & "&_f67r0=" & Me.Cells(row, 1).Value2 ' Detail line Function Code

            If detail_column_line_nbr > 0 Then
                If Me.Cells(row, detail_column_line_nbr).Value2 <> "" Then
                    request_parameters = request_parameters & "&_f79r0=" & Me.Cells(row, detail_column_line_nbr).Value2 ' JE line number, if there is one
                End If
            End If

            If detail_column_to_company > 0 Then
                If Me.Cells(row, detail_column_to_company).Value2 = "" Then
                    request_parameters = request_parameters & "&_f68r0=" & Me.Cells(row, detail_column_company).Value2
                Else
                    request_parameters = request_parameters & "&_f68r0=" & Me.Cells(row, detail_column_to_company).Value2 ' To Company, if specified; else Company from Header
                End If
            Else
                request_parameters = request_parameters & "&_f68r0=" & Me.Cells(row, detail_column_company).Value2
            End If

            request_parameters = request_parameters & "&_f69r0=" & Me.Cells(row, detail_column_acct_unit).Value2 ' Accounting Unit (cost center)
            request_parameters = request_parameters & "&_f70r0=" & Me.Cells(row, detail_column_account).Value2 ' GL Account

            If detail_column_sub_account > 0 Then
                If Me.Cells(row, detail_column_sub_account).Value2 <> "" Then request_parameters = request_parameters & "&_f71r0=" & Me.Cells(row, detail_column_sub_account).Value2 ' GL Sub Account, if specified
            End If

            If detail_column_activity > 0 Then
                If Me.Cells(row, detail_column_activity).Value2 <> "" Then request_parameters = request_parameters & "&_f73r0=" & Me.Cells(row, detail_column_activity).Value2 ' Activity code, if specified
            End If

            If detail_column_acct_category > 0 Then
                If Me.Cells(row, detail_column_acct_category).Value2 <> "" Then request_parameters = request_parameters & "&_f74r0=" & Me.Cells(row, detail_column_acct_category).Value2 ' Account Category code, if specified
            End If

            If detail_column_auto_rev > 0 Then
                If Me.Cells(row, detail_column_auto_rev).Value2 <> "" Then request_parameters = request_parameters & "&_f86r0=" & Me.Cells(row, detail_column_auto_rev).Value2
            End If

            request_parameters = request_parameters & "&_f75r0=" & Me.Cells(row, detail_column_tran_amount).Value2 ' Transaction Amount
            request_parameters = request_parameters & "&_f81r0=" & FilterForWeb(Left(Me.Cells(row, detail_column_description).Value2, 30)) ' 30 characters max field length
            If je.source_code <> "" Then request_parameters = request_parameters & "&_f89r0=" & je.source_code ' Source from Header, if specified

            If detail_column_reference > 0 Then
                If Me.Cells(row, detail_column_reference).Value2 <> "" Then request_parameters = request_parameters & "&_f88r0=" & FilterForWeb(Me.Cells(row, detail_column_reference).Value2)
            End If

            request_parameters = request_parameters & "&_OUT=XML&_INITDTL=TRUE" ' Send response in XML; bypass requiring an inquire before change

            SetXMLObject
            request_response = SendURL(request_parameters, "T")
            If Not g_oDom.LoadXML(request_response) Then
                If Me.Cells(row, 1).Value2 = "A" Then
                    Me.Cells(row, 1).Value2 = "C"
                    Me.Cells(row, detail_column_response).Value2 = "Loading error - check if line exists before adding again."
                Else
                    Me.Cells(row, detail_column_response).Value2 = "Loading error - check JE report to confirm change."
                End If
                Exit Sub
            End If

            For Each response_element In g_oDom.SelectNodes(search_string) ' Decide what to do with each text element based on the parent node's name
                Select Case response_element.ParentNode.nodeName
                    Case "Message"      ' This is the node name that returns the upload result
                        Me.Cells(row, detail_column_response).Value2 = Me.Cells(row, detail_column_response).Value2 & response_element.Text
                    Case "FldNbr"       ' This is the field number related to the response message
                        Me.Cells(row, detail_column_response).Value2 = Me.Cells(row, detail_column_response).Value2 & "(" & response_element.Text & ")"
                    Case "MsgNbr"       ' This return value refers to the message response - message number "000" is success
                        message_code = Val(response_element.Text)
                    Case "StatusNbr"    ' This is the status of the header upload. "001" = received
                        status = Val(response_element.Text)
                    Case "_f79r0"         ' This is the line number
                        If detail_column_line_nbr > 0 Then
                            Me.Cells(row, detail_column_line_nbr).Value2 = response_element.Text
                        End If
                End Select
            Next response_element

            If status = 1 And message_code = 0 Then ' note that status has been repurposed from server response ' If add/change/delete successful, delete the Function Code
                If detail_column_line_nbr > 0 Then
                    If Me.Cells(row, 1).Value = "D" Then Me.Cells(row, detail_column_line_nbr).Value2 = "deleted (" & Me.Cells(row, detail_column_line_nbr).Value2 & ")" ' and indicate successful delete
                End If
                Me.Cells(row, 1).Value = ""
            End If

        End If
    Next row
End Sub
Private Function ReportSheet() As Worksheet
On Error GoTo NoReport
    Set ReportSheet = ActiveWorkbook.Sheets("Report")
    Exit Function
NoReport:
    Set ReportSheet = ActiveSheet
End Function
Private Sub Worksheet_FollowHyperlink(ByVal Target As Hyperlink)
On Error GoTo error_handler
    Dim tip_text As String
    tip_text = Target.ScreenTip
    If Left(tip_text, 20) = "Journal Entry Report" Then
        tip_text = Replace(Replace(tip_text, "Journal Entry Report [", ""), "]", "")
        Dim params() As String
        params = Split(tip_text, ";")
        Sheets("Report").JournalEntryReport Co:=CInt(params(0)), Sys:=params(1), JeType:=params(2), _
            CtrlGrp:=CLng(params(3)), FY:=CInt(params(4)), Pd:=CInt(params(5))
    End If
    Exit Sub
error_handler:
    Target.Range.Worksheet.Activate
    Target.Delete
    MsgBox ("Error running JE report [" & tip_text & "]: " & Err.description)
End Sub
