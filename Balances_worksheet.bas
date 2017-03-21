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
Private Sub BalanceQuery_Click()
On Error GoTo error_handler
    Dim query_parameters As String
    Dim query_response As String
    Dim response_array() As String
    Dim data_range As Range

    If Not CheckUserAttributes() Then Login
    Set data_range = Me.Range("query_output")(2, 1)
    DeleteRows below_range:=data_range ' delete any previous outputs and reset UsedRange
    Range(Me.Range("query_output")(2, 1), Me.Range("query_output")(2, 19)).Clear ' clear first row without deleting table formula
    Me.Range("query_errors").EntireRow.Clear
    Me.Range("query_errors").Value = "Error messages go here:"

    query_parameters = "PROD=" & g_sProductLine & "&FILE=GLAMOUNTS&INDEX=GAMSET1&KEY=" & _
        FilterForWeb(Me.Range("query_company").Value & "=" & Me.Range("query_fy").Value & "=" & _
        Me.Range("query_acctunit").Value & "=" & Me.Range("query_account").Value) & "&FIELD=" & _
        FilterForWeb("COMPANY;ACCT-UNIT;ACCOUNT;SUB-ACCOUNT;CHART-DETAIL.ACCOUNT-DESC;FISCAL-YEAR;CYBAMT;CYPAMT1;CYPAMT2;CYPAMT3;CYPAMT4;CYPAMT5;CYPAMT6;CYPAMT7;CYPAMT8;CYPAMT9;CYPAMT10;CYPAMT11;CYPAMT12") & _
        "&OUT=XML&NEXT=FALSE&MAX=10000&keyUsage=PARAM"
    query_response = SendURL(query_parameters, "D")
    SetXMLObject ' Load page document into XML document object
    If Not g_oDom.LoadXML(query_response) Then
        Me.Cells(Me.Range("query_errors").row, Me.Cells(Me.Range("query_errors").row, 16384).End(xlToLeft).Column + 1).Value = "Could not load XML data from server."
        Exit Sub
    End If
    If g_oDom.DocumentElement.SelectSingleNode("/DME") Is Nothing Then ' do we have a /DME xml document?
        Me.Cells(Me.Range("query_errors").row, Me.Cells(Me.Range("query_errors").row, 16384).End(xlToLeft).Column + 1).Value = GetNodeText("/ERROR/MSG", 1) & " GLAMOUNTS Query" ' Error message from GLAMOUNTS query
        Exit Sub
    End If

    Application.ScreenUpdating = False ' Faster w/o updating screen for each cell - especially if output range is a table
    DmeToArray dme_array:=response_array
    ArrayToRange dme_array:=response_array, destination_range:=data_range
    Application.ScreenUpdating = True ' Revert to showing user the output

    FixObjects
    Exit Sub
error_handler:
    Application.ScreenUpdating = True
    Me.Cells(Me.Range("query_errors").row, Me.Cells(Me.Range("query_errors").row, 16384).End(xlToLeft).Column + 1).Value = "Error " & Err.Number & ": " & Err.description
    Resume Next
End Sub
Private Sub FixObjects() ' See Utilities/fix_ActiveX_objects.bas
    ' BalanceQuery
    Me.BalanceQuery.Left = 202.5
    Me.BalanceQuery.Width = 61.5
    Me.BalanceQuery.Height = 20.25
    Me.BalanceQuery.Top = 15.75
    Me.Shapes("BalanceQuery").ScaleHeight 1.25, msoFalse, msoScaleFromTopLeft
    Me.Shapes("BalanceQuery").ScaleHeight 0.8, msoFalse, msoScaleFromTopLeft
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
