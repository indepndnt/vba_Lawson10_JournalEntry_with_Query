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
Public Sub inGLTap()
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
    Dim iGlDesc As Integer ' Column of DESCRIPTION field (for hyperlink)
    Dim shtReport As Worksheet
    Dim iItems As Integer
    Dim sAttName As String
    Dim sAttText As String
    Dim bInclAp As Boolean
    Dim bExclChk As Boolean
    Dim bImgOpen As Boolean

    If Not CheckUserAttributes() Then Login
    bInclAp = Me.Range("incl_ap").Value
    bExclChk = Me.Range("excl_chk").Value
    bImgOpen = Me.Range("img_open").Value
    Me.Range("query_errors").EntireRow.Clear ' clear any previous error messages
    Me.Range("query_errors").Value = "Error messages go here:"
    Call fDeleteFrom(Me.Range("query_output")(2, 1)) ' delete any previous outputs and reset UsedRange

    Set shtReport = ReportSheet
    Set rngOut = Me.Range("query_output", Me.Range("query_output").End(xlToRight)) ' Build query field list from worksheet column heading names
    iColumns = rngOut.Columns.Count
    sFields = ""
    For Each rCell In rngOut
        If sFields <> "" Then sFields = sFields & ";"
        sFields = sFields & rCell.Value
        If rCell.Value = "DESCRIPTION" Then iGlDesc = rCell.Column
        If rCell.Value = "OBJ-ID" Then iGlObjId = rCell.Column
        If rCell.Value = "APDISTRIB.API-OBJ-ID" Then iApObjId = rCell.Column
    Next rCell
    If (iGlObjId = 0 Or iApObjId = 0) And bInclAp Then
        Me.Range("incl_ap").Value = False
        NonFatal ("The 'OBJ-ID' or 'APDISTRIB.API-OBJ-ID' columns are missing.")
    End If

    Set rngOut = Me.Range("query_output")(2, 1)
    If Me.Range("max_records").Value > 10000 Then Me.Range("max_records").Value = 10000

    s = "PROD=" & g_sProductLine & "&FILE=GLTRANS&INDEX=GLTSET3" ' Table GLTRANS, criteria set GLTSET3: key = co=account=subacct=acct-unit=fy=pd
    s = s & "&KEY=" & FilterForWeb(Me.Range("query_company").Value & "=" & Me.Range("query_account").Value & "=0=" & _
        Me.Range("query_acctunit").Value & "=" & Me.Range("query_fy").Value & "=" & Me.Range("query_period").Value)
    s = s & "&FIELD=" & FilterForWeb(sFields) & "&OUT=XML&NEXT=FALSE&MAX=" & Me.Range("max_records").Value & "&keyUsage=PARAM"
    s = SendURL(s, "D")
    SetXMLObject ' Load IE page document into XML document object
    If Not g_oDom.LoadXML(s) Then Exit Sub
    If Not inXmlDme Then ' do we have a /DME xml document?
        Call NonFatal(inXmlData("/ERROR/MSG", 1), "GLTRANS Query") ' Error message from GLTRANS query
        Exit Sub
    End If

    iCount = Val(g_oDom.DocumentElement.SelectSingleNode("//RECORDS").Attributes.getNamedItem("count").Text) ' Get count of records returned
    If iCount = 0 Then
        rngOut(2, 1).Value = "No results returned."
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Dim qArray() As String
    Call inQueryArray(qArray)
    Call inArrayToRange(qArray, rngOut)
    If Not shtReport Is ActiveSheet Then
        For row = 1 To iCount
            rngOut(row, iGlDesc).Hyperlinks.Add Anchor:=rngOut(row, iGlDesc), Address:="", ScreenTip:="Journal Entry Report", _
                SubAddress:="'" & shtReport.Name & "'!" & shtReport.Range("A12").Address
        Next row
    End If
    Application.ScreenUpdating = True

    If bInclAp Then
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
                            iItems = Val(g_oDom.DocumentElement.SelectSingleNode("/ACTION/LIST").Attributes.getNamedItem("numRecords").Text) ' number of matching attachments
                            If iItems = 0 Then
                                rngOut(row, iColumns + 1).Value = "no images"
                            Else
                                col = iColumns
                                For i = 1 To iItems
                                    sAttName = inXmlData("//ATTACHMENT", i, "ATTACHMENT-NAME") ' //ATTACHMENT/ATTACHMENT-NAME ("Check Image" or "Invoice Image")
                                    sAttText = inXmlData("//ATTACHMENT", i, "ATTACHMENT-TEXT") ' //ATTACHMENT/ATTACHMENT-TEXT (image URL)
                                    If Not bExclChk Or Left(sAttName, 5) <> "Check" Then
                                        col = col + 1
                                        rngOut(row, col).Value = sAttName & ": " & sAttText
                                        ActiveSheet.Hyperlinks.Add Anchor:=rngOut(row, col), Address:=sAttText
                                        If bImgOpen Then rngOut(row, col).Hyperlinks(1).Follow ' Open image if flag set
                                    End If
                                Next i
                            End If
                        End If
                    End If
                End If
            End If
        Next row
    End If

    Exit Sub
errHandler:
    Call NonFatal("Error " & Err.Number & ": " & Err.Description, "Row " & row + rngOut.row)
    Resume Next
End Sub
Private Function ReportSheet() As Worksheet
    On Error GoTo NoReport
    Set ReportSheet = ActiveWorkbook.Sheets("Report")
    Exit Function
NoReport:
    Set ReportSheet = ActiveSheet
End Function
Private Sub Worksheet_FollowHyperlink(ByVal Target As Hyperlink)
    If Target.ScreenTip = "Journal Entry Report" Then
        Dim rng As Range
        Dim col(6) As Integer
        For Each rng In Me.Range("query_output", Me.Range("query_output").End(xlToRight))
            Select Case rng.Value2
                Case "COMPANY"
                    col(0) = rng.Column
                Case "SYSTEM"
                    col(1) = rng.Column
                Case "JE-TYPE"
                    col(2) = rng.Column
                Case "GLCONTROL.CONTROL-GROUP"
                    col(3) = rng.Column
                Case "FISCAL-YEAR"
                    col(4) = rng.Column
                Case "ACCT-PERIOD"
                    col(5) = rng.Column
            End Select
        Next rng
        Dim sht As Worksheet
        Set sht = ThisWorkbook.Sheets(Replace(Split(Target.SubAddress, "!")(0), "'", ""))
        Set rng = Target.Range.Worksheet.Cells(Target.Range.row, 1)
        Sheets("Report").inGL240 Co:=rng(1, col(0)).Value2, Sys:=rng(1, col(1)).Value2, JeType:=rng(1, col(2)).Value2, _
            CtrlGrp:=rng(1, col(3)).Value2, FY:=rng(1, col(4)).Value2, Pd:=rng(1, col(5)).Value2
    End If
End Sub
Private Sub NonFatal(ByVal sMsg As String, Optional ByVal sRange As String = "")
    Dim iColumn As Long
    Dim sAppend As String
    iColumn = Me.Cells(Me.Range("query_errors").row, 16384).End(xlToLeft).Column + 1
    If sRange <> "" Then
        sAppend = " (" & sRange & ")"
    Else
        sAppend = ""
    End If
    Me.Cells(Me.Range("query_errors").row, iColumn).Value = sMsg & sAppend
End Sub
