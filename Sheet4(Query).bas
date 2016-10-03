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
Public Sub inGLInvoices() ' Used to run only inGLTap - point 'Query' button on the Query tab to here.
On Error GoTo errHandler
    If Not CheckUserAttributes() Then Login
    If Not inGLTap() Then
        Call NonFatal("Query Error")
    End If
    Exit Sub
errHandler:
    MsgBox ("Error: " & Err.Number & ":" & Err.Description)
End Sub
Public Function inGLTap() As Boolean
On Error GoTo errHandler
    Const urlDrill As String = "/servlet/Router/Drill/erp?"  ' Infor server path for drill
    Const urlAttach As String = "/lawson-ios/action/ListAttachments?" ' Infor server path for Attachment List
    Dim xmlNode As MSXML2.IXMLDOMNode
    Dim s As String
    Dim i As Integer
    Dim rngOut As Range ' Output table range
    Dim row As Long ' Output table row
    Dim col As Long
    Dim iCount As Long
    Dim iColumns As Integer
    Dim sFields As String
    Dim rCell As Range
    Dim iGlObjId As Integer ' Column of GLTRANS OBJ_ID field
    Dim iApObjId As Integer ' Column of APINVOICE OBJ_ID field
    Dim iItems As Integer
    Dim sAttName As String
    Dim sAttText As String

    inGLTap = False ' Defaults if no result
    Sheet4.Range("6:6").Clear ' clear any previous error messages
    Sheet4.Range("A6").Value = "Error messages go here:"
    Call fDeleteFrom(Sheet4.Range("A9")) ' delete any previous outputs and reset UsedRange

    Set rngOut = Sheet4.Range("A8", Range("A8").End(xlToRight)) ' Build query field list from worksheet column heading names
    iColumns = rngOut.Columns.Count
    sFields = ""
    For Each rCell In rngOut
        If sFields <> "" Then sFields = sFields & ";"
        sFields = sFields & rCell.Value
        If rCell.Value = "OBJ-ID" Then iGlObjId = rCell.Column
        If rCell.Value = "APDISTRIB.API-OBJ-ID" Then iApObjId = rCell.Column
    Next rCell
    If (iGlObjId = 0 Or iApObjId = 0) And Range("Q1").Value Then
        Range("Q1").Value = False
        NonFatal ("The 'OBJ-ID' or 'APDISTRIB.API-OBJ-ID' columns are missing.")
    End If

    Set rngOut = Sheet4.Range("A9")

    s = "PROD=" & g_sProductLine ' Start building POST data string with Product Line
    s = s & "&FILE=GLTRANS&INDEX=GLTSET3" ' ' Table GLTRANS, criteria set GLTSET3: key = co=account=subacct=acct-unit=fy=pd
    s = s & "&KEY=" & FilterForWeb(Sheet4.Range("D1").Value & "=" & Sheet4.Range("D3").Value & "=0=" & Sheet4.Range("D2").Value & "=" & Sheet4.Range("D4").Value & "=" & Sheet4.Range("D5").Value)
    s = s & "&FIELD=" & FilterForWeb(sFields)
    s = s & "&OUT=XML&NEXT=FALSE" ' NEXT=FALSE means don't give me the RELOAD string.
    s = s & "&MAX=" & Sheet4.Range("Q3").Value & "&keyUsage=PARAM" ' Give me up to (value in cell Q3) records.
    s = SendURL(s, "D")
    SetXMLObject ' Load IE page document into XML document object
    If Not g_oDom.LoadXML(s) Then
        inGLTap = False ' If we couldn't load g_oDom with the Lawson output then exit with an error - there's no data.
        Exit Function
    End If
    If Not inXmlDme Then ' do we have a /DME xml document?
        Call NonFatal(inXmlData("/ERROR/MSG", 1), "GLTRANS Query") ' Error message from GLTRANS query
        inGLTap = True
        Exit Function
    End If

    iCount = Val(inXmlAttribVal()) ' Get count of records returned

    For row = 1 To iCount ' Copy GL query data into worksheet
        For col = 1 To iColumns
            Call inCellType(rngOut, row, col)
        Next col
    Next row

    If Sheet4.Range("Q1").Value Then
        For row = 1 To iCount ' For each result with an API OBJ-ID, try to get image URLs
            If rngOut(row, iApObjId).Value > 0 Then
                s = "_PDL=" & g_sProductLine
                s = s & "&_TYP=OV&_IN=APDSET5&_RID=AP-APD-V-0002&_SYS=GL&K3=1"
                s = s & "&K1=" & rngOut(row, iGlObjId).Value ' GLTRANS Object ID
                s = s & "&K2=" & rngOut(row, iApObjId).Value ' APINVOICE Object ID
                s = s & "&K3=1&keyUsage=PARAM&_RECSTOGET=1"
                s = SendURL(g_sServer & urlDrill & s, "X")
                If Not g_oDom.LoadXML(s) Then
                    Call NonFatal("Could not load Attachment List: ", "Row " & row + rngOut.row)
                Else
                    If Not inXmlDme("/IDARETURN") Then ' check - is this IDARETURN, or error
                        Call NonFatal(inXmlData("/ERROR/MSG", 1), "Drill Query (row " & row + rngOut.row & ")") ' Error message from Drill query
                    Else
                        s = "dataArea=" & g_sProductLine
                        s = s & "&attachmentType=I&drillType=URL&objName=Invoice URL Attachment&attachmentCategory=U&indexName=APISET1&fileName=APINVOICE"
                        s = s & "&K1=" & rngOut(row, 1).Value ' Company
                        s = s & "&K2=" & inXmlData("//LINE/COLS/COL", 1) ' Vendor
                        s = s & "&K3=" & inXmlData("//LINE/COLS/COL", 2) ' Invoice #
                        s = s & "&K4=0&K5=0&outType=XML"
                        s = SendURL(g_sServer & urlAttach & s, "X")
                        If Not g_oDom.LoadXML(s) Then
                            Call NonFatal("Could not load API Attachments: ", "Row " & row + rngOut.row)
                        Else
                            iItems = Val(inXmlAttribVal("/ACTION/LIST", "numRecords")) ' number of matching attachments
                            If iItems = 0 Then
                                rngOut(row, iColumns + 1).Value = "no images"
                            Else
                                For i = 1 To iItems
                                    sAttName = inXmlData("//ATTACHMENT", i, "ATTACHMENT-NAME") ' //ATTACHMENT/ATTACHMENT-NAME ("Check Image" or "Invoice Image")
                                    sAttText = inXmlData("//ATTACHMENT", i, "ATTACHMENT-TEXT") ' //ATTACHMENT/ATTACHMENT-TEXT (image URL)
                                    rngOut(row, iColumns + i).Value = sAttName & ": " & sAttText
                                    ActiveSheet.Hyperlinks.Add Anchor:=rngOut(row, iColumns + i), Address:=sAttText
                                    If Sheet4.Range("Q2") And Left(sAttName, 7) = "Invoice" Then rngOut(row, iColumns + i).Hyperlinks(1).Follow ' Open image if flag set
                                Next i
                            End If
                        End If
                    End If
                End If
            End If
        Next row
    End If
    
    inGLTap = True
    Exit Function
errHandler:
    Call NonFatal("Error " & Err.Number & ": " & Err.Description, "Row " & row + rngOut.row)
    Resume Next
End Function
Private Sub NonFatal(ByVal sMsg As String, Optional ByVal sRange As String = "")
    Dim iColumn As Long
    Dim sAppend As String
    iColumn = Range("XFD6").End(xlToLeft).Column + 1
    If sRange <> "" Then
        sAppend = " (" & sRange & ")"
    Else
        sAppend = ""
    End If
    Cells(6, iColumn).Value = sMsg & sAppend
End Sub
