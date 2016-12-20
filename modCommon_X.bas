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
Public Function fDeleteFrom(ByVal rng As Range) As Boolean
    On Error GoTo errHandler

    Dim sht As Worksheet
    Set sht = ActiveSheet

    Dim rSheet As Worksheet
    Set rSheet = rng.Worksheet

    Dim lUsedRow As Long
    lUsedRow = rSheet.UsedRange.Rows.Count ' last row in UsedRange

    Dim lRangeRow As Long
    lRangeRow = rng.row

    fDeleteFrom = False
    rSheet.Activate ' activate destination worksheet to faciliate resetting UsedRange
    If lUsedRow > lRangeRow Then ' if our destination row is after UsedRange, we'd end up deleting the last row of UsedRange (e.g., rows("10:2").delete)
        rng.Worksheet.Rows(lRangeRow & ":" & lUsedRow).Delete
    End If
    Application.ActiveSheet.UsedRange ' Reset the Used Range to control file size
    If rSheet.UsedRange.Rows.Count < lRangeRow + 2 Then fDeleteFrom = True
    sht.Activate
    Exit Function
errHandler:
    Debug.Print "Function fDeleteFrom error " & Err.Number & " (" & Err.Description & ")"
    Resume Next
End Function
Public Function inXmlDme(Optional ByVal sRoot As String = "/DME") As Boolean
    Dim xmlNode As MSXML2.IXMLDOMNode
    Set xmlNode = g_oDom.DocumentElement.SelectSingleNode(sRoot)
    If xmlNode Is Nothing Then
        inXmlDme = False
    Else
        inXmlDme = True
    End If
End Function
Public Sub inCellType(ByVal rBase As Range, ByVal iRow As Integer, ByVal iCol As Integer)
    Dim sType As String
    Dim s As String
    sType = inXmlAttribVal("//COLUMN[" & iCol & "]", "type")
    s = inQueryCell(iRow, iCol)
    Select Case sType
Option Explicit
Public Function fDeleteFrom(ByVal rng As Range) As Boolean
    On Error GoTo errHandler

    Dim sht As Worksheet
    Set sht = ActiveSheet

    Dim rSheet As Worksheet
    Set rSheet = rng.Worksheet

    Dim lUsedRow As Long
    lUsedRow = rSheet.UsedRange.Rows.Count ' last row in UsedRange

    Dim lRangeRow As Long
    lRangeRow = rng.row

    fDeleteFrom = False
    rSheet.Activate ' activate destination worksheet to faciliate resetting UsedRange
    If lUsedRow > lRangeRow Then ' if our destination row is after UsedRange, we'd end up deleting the last row of UsedRange (e.g., rows("10:2").delete)
        rng.Worksheet.Rows(lRangeRow & ":" & lUsedRow).Delete
    End If
    Application.ActiveSheet.UsedRange ' Reset the Used Range to control file size
    If rSheet.UsedRange.Rows.Count < lRangeRow + 2 Then fDeleteFrom = True
    sht.Activate
    Exit Function
errHandler:
    Debug.Print "Function fDeleteFrom error " & Err.Number & " (" & Err.Description & ")"
    Resume Next
End Function
Public Function inXmlDme(Optional ByVal sRoot As String = "/DME") As Boolean
    Dim xmlNode As MSXML2.IXMLDOMNode
    Set xmlNode = g_oDom.DocumentElement.SelectSingleNode(sRoot)
    If xmlNode Is Nothing Then
        inXmlDme = False
    Else
        inXmlDme = True
    End If
End Function
Public Sub inQueryArray(ByRef sArr() As String)
    Dim iRec As Integer
    Dim iCol As Integer
    Dim r As Integer
    Dim c As Integer
    Dim s As String
    Dim rNode As MSXML2.IXMLDOMNode
    Dim cNode As MSXML2.IXMLDOMNode
    Dim rNodes As MSXML2.IXMLDOMNodeList
    Dim cNodes As MSXML2.IXMLDOMNodeList
    Set rNodes = g_oDom.SelectNodes("/DME/RECORDS/RECORD/COLS")
    iRec = rNodes.Length
Option Explicit
Public Function fDeleteFrom(ByVal rng As Range) As Boolean
    On Error GoTo errHandler

    Dim sht As Worksheet
    Set sht = ActiveSheet

    Dim rSheet As Worksheet
    Set rSheet = rng.Worksheet

    Dim lUsedRow As Long
    lUsedRow = rSheet.UsedRange.Rows.Count ' last row in UsedRange

    Dim lRangeRow As Long
    lRangeRow = rng.row

    fDeleteFrom = False
    rSheet.Activate ' activate destination worksheet to faciliate resetting UsedRange
    If lUsedRow > lRangeRow Then ' if our destination row is after UsedRange, we'd end up deleting the last row of UsedRange (e.g., rows("10:2").delete)
        rng.Worksheet.Rows(lRangeRow & ":" & lUsedRow).Delete
    End If
    Application.ActiveSheet.UsedRange ' Reset the Used Range to control file size
    If rSheet.UsedRange.Rows.Count < lRangeRow + 2 Then fDeleteFrom = True
    sht.Activate
    Exit Function
errHandler:
    Debug.Print "Function fDeleteFrom error " & Err.Number & " (" & Err.Description & ")"
    Resume Next
End Function
Public Function inXmlDme(Optional ByVal sRoot As String = "/DME") As Boolean
    Dim xmlNode As MSXML2.IXMLDOMNode
    Set xmlNode = g_oDom.DocumentElement.SelectSingleNode(sRoot)
    If xmlNode Is Nothing Then
        inXmlDme = False
    Else
        inXmlDme = True
    End If
End Function
Public Sub inQueryArray(ByRef sArr() As String)
    Dim iRec As Integer
    Dim iCol As Integer
    Dim r As Integer
    Dim c As Integer
    Dim s As String
    Dim rNode As MSXML2.IXMLDOMNode
    Dim cNode As MSXML2.IXMLDOMNode
    Dim rNodes As MSXML2.IXMLDOMNodeList
    Dim cNodes As MSXML2.IXMLDOMNodeList
    Set rNodes = g_oDom.SelectNodes("/DME/RECORDS/RECORD/COLS")
    iRec = rNodes.Length
    iCol = g_oDom.SelectNodes("/DME/COLUMNS/COLUMN").Length
    ReDim sArr(iRec, iCol)
    r = 0
    For Each rNode In rNodes
        Set cNodes = rNode.SelectNodes("COL")
        c = 0
        For Each cNode In cNodes
            s = cNode.FirstChild.Text
            If Len(s) > 0 Then
                sArr(r, c) = s
            End If
            c = c + 1
        Next cNode
        r = r + 1
    Next rNode
End Sub
Public Sub inArrayToRange(ByRef sArr() As String, ByVal rDest As Range)
    Dim nodes As MSXML2.IXMLDOMNodeList
    Dim c As Integer
    Dim col As Range
    Dim fmt As XlColumnDataType
    Dim sType As String
    Set rDest = rDest.Resize(UBound(sArr, 1), UBound(sArr, 2))
    rDest.Value = sArr
    Set nodes = g_oDom.SelectNodes("/DME/COLUMNS/COLUMN")
    For c = 1 To nodes.Length
        Set col = rDest.Columns(c)
        sType = nodes(c - 1).Attributes.getNamedItem("type").Text
        Select Case sType
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
Public Function inXmlData(ByVal sField As String, ByVal iOrdinal As Integer, Optional ByVal sChildNode = "/") As String
    Dim xmlNode As MSXML2.IXMLDOMNode
    Dim XmlNodeList As MSXML2.IXMLDOMNodeList
    Dim sChild As String
    If sChildNode = "/" Then
        sChild = ""
    Else
        sChild = "/" & sChildNode
    End If
    Set XmlNodeList = g_oDom.DocumentElement.SelectNodes(sField)
    If XmlNodeList.Length > 0 And XmlNodeList.Length >= iOrdinal Then
        Set xmlNode = g_oDom.DocumentElement.SelectSingleNode(sField & "[" & iOrdinal & "]" & sChild)
        inXmlData = xmlNode.FirstChild.Text
    Else
        inXmlData = ""
    End If
End Function
