' VBA Lawson Excel Tools
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
' Home is https://github.com/indepndnt/vba-Lawson-Excel-Tools
'
Option Explicit
Private Sub GL_Query_Click()
On Error GoTo error_handler
    Const drill_url As String = "/servlet/Router/Drill/erp?"  ' Infor server path for drill
    Const attachment_url As String = "/lawson-ios/action/ListAttachments?" ' Infor server path for Attachment List
    Dim query_parameters As String
    Dim query_response As String
    Dim i As Integer
    Dim data_range As Range ' Output table range
    Dim row As Long ' Output table row
    Dim col As Long
    Dim record_count As Long
    Dim column_count As Integer
    Dim query_fields As String
    Dim cell_range As Range
    Dim gl_id_column As Integer ' Column of GLTRANS OBJ_ID field
    Dim ap_id_column As Integer ' Column of APINVOICE OBJ_ID field
    Dim description_column As Integer ' Column of DESCRIPTION field (for hyperlink)
    Dim company_col As Integer ' Column of COMPANY field (for hyperlink)
    Dim system_col As Integer ' Column of SYSTEM field (for hyperlink)
    Dim je_type_col As Integer ' Column of JE-TYPE field (for hyperlink)
    Dim control_group_col As Integer ' Column of CONTROL-GROUP field (for hyperlink)
    Dim fiscal_year_col As Integer ' Column of FISCAL-YEAR field (for hyperlink)
    Dim acct_period_col As Integer ' Column of ACCT-PERIOD field (for hyperlink)
    Dim attachment_count As Integer
    Dim attachment_name As String
    Dim attachment_text As String
    Dim response_array() As String
    Dim can_link As Boolean

    If Not CheckUserAttributes() Then Login
    Me.Range("query_errors").EntireRow.Clear ' clear any previous error messages
    Me.Range("query_errors").Value = "Error messages go here:"
    DeleteRows below_range:=Me.Range("query_output")(2, 1) ' delete any previous outputs and reset UsedRange

    Set data_range = Me.Range("query_output", Me.Range("query_output").End(xlToRight)) ' Build query field list from worksheet column heading names
    column_count = data_range.Columns.Count
    query_fields = ""
    For Each cell_range In data_range
        If query_fields <> "" Then query_fields = query_fields & ";"
        query_fields = query_fields & cell_range.Value
        If cell_range.Value = "DESCRIPTION" Then description_column = cell_range.Column
        If cell_range.Value = "COMPANY" Then company_col = cell_range.Column
        If cell_range.Value = "SYSTEM" Then system_col = cell_range.Column
        If cell_range.Value = "JE-TYPE" Then je_type_col = cell_range.Column
        If cell_range.Value = "GLCONTROL.CONTROL-GROUP" Then control_group_col = cell_range.Column
        If cell_range.Value = "FISCAL-YEAR" Then fiscal_year_col = cell_range.Column
        If cell_range.Value = "ACCT-PERIOD" Then acct_period_col = cell_range.Column
        If cell_range.Value = "OBJ-ID" Then gl_id_column = cell_range.Column
        If cell_range.Value = "APDISTRIB.API-OBJ-ID" Then ap_id_column = cell_range.Column
    Next cell_range
    If (gl_id_column = 0 Or ap_id_column = 0) And Me.include_images.Value Then
        Me.include_images.Value = False
        NonFatal ("The 'OBJ-ID' or 'APDISTRIB.API-OBJ-ID' columns are missing.")
    End If

    Set data_range = Me.Range("query_output")(2, 1)
    If Me.Range("max_records").Value > 10000 Then Me.Range("max_records").Value = 10000

    query_parameters = "PROD=" & g_sProductLine & "&FILE=GLTRANS&INDEX=GLTSET3&KEY=" & _
        FilterForWeb(Me.Range("query_company").Value & "=" & Me.Range("query_account").Value & "=0->9999=" & _
        Me.Range("query_acctunit").Value & "=" & Me.Range("query_fy").Value & "=" & Me.Range("query_period").Value) & _
        "&FIELD=" & FilterForWeb(query_fields) & "&OUT=XML&NEXT=FALSE&MAX=" & _
        Me.Range("max_records").Value & "&keyUsage=PARAM"
    query_response = SendURL(query_parameters, "D")
    SetXMLObject ' Load page document into XML document object
    If Not g_oDom.LoadXML(query_response) Then Exit Sub
    If g_oDom.DocumentElement.SelectSingleNode("/DME") Is Nothing Then ' do we have a /DME xml document?
        Call NonFatal(GetNodeText("/ERROR/MSG", 1), "GLTRANS Query") ' Error message from GLTRANS query
        Exit Sub
    End If

    record_count = Val(g_oDom.DocumentElement.SelectSingleNode("//RECORDS").Attributes.getNamedItem("count").Text) ' Get count of records returned
    If record_count = 0 Then
        data_range(2, 1).Value = "No results returned."
        Exit Sub
    End If

    can_link = company_col > 0 And system_col > 0 And je_type_col > 0 And control_group_col > 0 And fiscal_year_col > 0 And acct_period_col > 0 _
        And Not ReportSheet Is ActiveSheet

    Application.ScreenUpdating = False
    DmeToArray dme_array:=response_array
    ArrayToRange dme_array:=response_array, destination_range:=data_range
    If can_link Then
        If description_column = 0 Then description_column = 1
        For row = 1 To record_count
            data_range(row, description_column).Hyperlinks.Add Anchor:=data_range(row, description_column), Address:="", SubAddress:="Report!$A$12", _
                ScreenTip:="Journal Entry Report [" & Join(Array(Trim(response_array(row - 1, company_col - 1)), response_array(row - 1, system_col - 1), _
                response_array(row - 1, je_type_col - 1), Trim(response_array(row - 1, control_group_col - 1)), _
                response_array(row - 1, fiscal_year_col - 1), response_array(row - 1, acct_period_col - 1)), ";") & "]"
        Next row
    End If
    Application.ScreenUpdating = True

    If Me.include_images.Value Then
        For row = 1 To record_count ' For each result with an API OBJ-ID, try to get image URLs
            If data_range(row, ap_id_column).Value > 0 Then
                query_parameters = "_PDL=" & g_sProductLine & "&_TYP=OV&_IN=APDSET5&_RID=AP-APD-V-0002&_SYS=GL&K3=1" & _
                    "&K1=" & data_range(row, gl_id_column).Value & "&K2=" & data_range(row, ap_id_column).Value & _
                    "&K3=1&keyUsage=PARAM&_RECSTOGET=1"
                query_response = SendURL(g_sServer & drill_url & query_parameters, "X")
                If Not g_oDom.LoadXML(query_response) Then
                    Call NonFatal("Could not load Attachment List: ", "Row " & row + data_range.row)
                Else
                    If g_oDom.DocumentElement.SelectSingleNode("/IDARETURN") Is Nothing Then ' check - is this IDARETURN, or error
                        Call NonFatal(GetNodeText("/ERROR/MSG", 1), "Drill Query (row " & row + data_range.row & ")") ' Error message from Drill query
                    Else
                        query_parameters = "dataArea=" & g_sProductLine & "&attachmentType=I&drillType=URL&objName=Invoice URL Attachment&attachmentCategory=U&indexName=APISET1&fileName=APINVOICE" & _
                            "&K1=" & data_range(row, 1).Value & "&K2=" & GetNodeText("//LINE/COLS/COL", 1) & _
                            "&K3=" & GetNodeText("//LINE/COLS/COL", 2) & "&K4=0&K5=0&outType=XML"
                        query_response = SendURL(g_sServer & attachment_url & query_parameters, "X")
                        If Not g_oDom.LoadXML(query_response) Then
                            Call NonFatal("Could not load API Attachments: ", "Row " & row + data_range.row)
                        Else
                            attachment_count = Val(g_oDom.DocumentElement.SelectSingleNode("/ACTION/LIST").Attributes.getNamedItem("numRecords").Text) ' number of matching attachments
                            If attachment_count = 0 Then
                                data_range(row, column_count + 1).Value = "no images"
                            Else
                                col = column_count
                                For i = 1 To attachment_count
                                    attachment_name = GetNodeText("//ATTACHMENT", i, "ATTACHMENT-NAME") ' //ATTACHMENT/ATTACHMENT-NAME ("Check Image" or "Invoice Image")
                                    attachment_text = GetNodeText("//ATTACHMENT", i, "ATTACHMENT-TEXT") ' //ATTACHMENT/ATTACHMENT-TEXT (image URL)
                                    If Not Me.exclude_checks.Value Or Left(attachment_name, 5) <> "Check" Then
                                        col = col + 1
                                        data_range(row, col).Value = attachment_name & ": " & attachment_text
                                        ActiveSheet.Hyperlinks.Add Anchor:=data_range(row, col), Address:=attachment_text
                                        If Me.open_images.Value Then data_range(row, col).Hyperlinks(1).Follow ' Open image if flag set
                                    End If
                                Next i
                            End If
                        End If
                    End If
                End If
            End If
        Next row
    End If

    FixObjects ' When ActiveX controls attack!
    Exit Sub
error_handler:
    Application.ScreenUpdating = True
    Call NonFatal("Error " & Err.Number & ": " & Err.description, "Row " & row + data_range.row)
    Resume Next
End Sub
Private Sub FixObjects() ' See Utilities/fix_ActiveX_objects.bas
    ' GL_Query
    Me.GL_Query.Left = 519.75
    Me.GL_Query.Width = 60.75
    Me.GL_Query.Height = 24.75
    Me.GL_Query.Top = 16.5
    Me.Shapes("GL_Query").ScaleHeight 1.25, msoFalse, msoScaleFromTopLeft
    Me.Shapes("GL_Query").ScaleHeight 0.8, msoFalse, msoScaleFromTopLeft

    ' include_images
    Me.include_images.Left = 594.75
    Me.include_images.Width = 105
    Me.include_images.Height = 15.75
    Me.include_images.Top = 0.75
    Me.Shapes("include_images").ScaleHeight 1.25, msoFalse, msoScaleFromTopLeft
    Me.Shapes("include_images").ScaleHeight 0.8, msoFalse, msoScaleFromTopLeft

    ' exclude_checks
    Me.exclude_checks.Left = 594.75
    Me.exclude_checks.Width = 105
    Me.exclude_checks.Height = 15.75
    Me.exclude_checks.Top = 18
    Me.Shapes("exclude_checks").ScaleHeight 1.25, msoFalse, msoScaleFromTopLeft
    Me.Shapes("exclude_checks").ScaleHeight 0.8, msoFalse, msoScaleFromTopLeft

    ' open_images
    Me.open_images.Left = 594.75
    Me.open_images.Width = 105
    Me.open_images.Height = 15.75
    Me.open_images.Top = 34.5
    Me.Shapes("open_images").ScaleHeight 1.25, msoFalse, msoScaleFromTopLeft
    Me.Shapes("open_images").ScaleHeight 0.8, msoFalse, msoScaleFromTopLeft
End Sub
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
Private Sub ArrayToRange(ByRef dme_array() As String, ByVal destination_range As Range)
    Dim nodes As MSXML2.IXMLDOMNodeList
    Dim c As Integer
    Dim col As Range
    Dim fmt As XlColumnDataType
    Dim column_data_type As String
    Set destination_range = destination_range.Resize(UBound(dme_array, 1), UBound(dme_array, 2))
    destination_range.Value = dme_array
    Set nodes = g_oDom.SelectNodes("/DME/COLUMNS/COLUMN")
    For c = 1 To nodes.Length
        Set col = destination_range.Columns(c)
        column_data_type = nodes(c - 1).Attributes.getNamedItem("type").Text
        Select Case column_data_type
            Case "BCD" ' "Binary Coded Decimal" is the format used with currency fields. It has a trailing minus for negative numbers.
                col.NumberFormat = "#,##0.00_);[Red](#,##0.00)"
                fmt = xlGeneralFormat
            Case "NUMERIC" ' This is the type for numeric codes (e.g. company code or fiscal year or object id)
                col.NumberFormat = "General"
                fmt = xlGeneralFormat
            Case "ALPHA", "ALPHALC" ' Text or alphanumeric fields (ALPHA is case insensetive and ALPHALC includes LowerCase)
                col.NumberFormat = "@"
                fmt = xlTextFormat
            Case "YYYYMMDD" ' (date)
                col.NumberFormat = "[$-409]d-mmm-yyyy;@"
                fmt = xlMDYFormat
        End Select
        If WorksheetFunction.CountA(col) > 0 Then
            col.TextToColumns Destination:=col, DataType:=xlDelimited, FieldInfo:=Array(1, fmt), TrailingMinusNumbers:=True
        End If
    Next c
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
    NonFatal error_message:="Error running JE report.", error_range:=tip_text
End Sub
Private Function ReportSheet() As Worksheet
On Error GoTo NoReport
    Set ReportSheet = ActiveWorkbook.Sheets("Report")
    Exit Function
NoReport:
    Set ReportSheet = ActiveSheet
End Function
Private Sub NonFatal(ByVal error_message As String, Optional ByVal error_range As String = "")
    Dim next_error_column As Long
    next_error_column = Me.Cells(Me.Range("query_errors").row, 16384).End(xlToLeft).Column + 1
    If error_range <> "" Then
        error_message = error_message & " (" & error_range & ")"
    End If
    Me.Cells(Me.Range("query_errors").row, next_error_column).Value = error_message
End Sub
