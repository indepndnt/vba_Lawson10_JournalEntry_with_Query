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
        Case "BCD" ' "Binary Coded Decimal" is the format used with currency fields. It has a trailing minus for negative numbers.
            s = Replace(s, " ", "")
            If s Like "*-" Then s = "-" & Replace(s, "-", "")
            rBase(iRow, iCol).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
            rBase(iRow, iCol).Value = Val(s)
        Case "NUMERIC" ' This is the type for numeric codes (e.g. company code or fiscal year or object id)
            rBase(iRow, iCol).NumberFormat = "General"
            rBase(iRow, iCol).Value = s
        Case "ALPHA", "ALPHALC" ' Text or alphanumeric fields (ALPHA is case insensetive and ALPHALC includes LowerCase)
            rBase(iRow, iCol).NumberFormat = "@"
            If Len(s) > 0 Then rBase(iRow, iCol).Value = s
        Case "YYYYMMDD" ' (date)
            rBase(iRow, iCol).NumberFormat = "[$-409]d-mmm-yyyy;@"
            rBase(iRow, iCol).Value = s
        Case Else
            rBase(iRow, iCol).Value = s
    End Select
End Sub
Public Function inQueryCell(ByVal iRow As Integer, ByVal iColumn As Integer, Optional sField As String = "//RECORD", Optional sChild As String = "/COLS/COL") As String
    inQueryCell = inXmlData(sField, iRow, sChild & "[" & iColumn & "]")
End Function
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
Public Function inXmlAttribVal(Optional ByVal sField As String = "//RECORDS", Optional ByVal sAttrib As String = "count") As String
    Dim xmlNode As MSXML2.IXMLDOMNode
    Set xmlNode = g_oDom.DocumentElement.SelectSingleNode(sField)
    If xmlNode Is Nothing Then
        inXmlAttribVal = False
        Exit Function
    Else
        inXmlAttribVal = xmlNode.Attributes.getNamedItem(sAttrib).Text
    End If
End Function
