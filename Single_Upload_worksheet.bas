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
Private Sub JournalEntryReport_Click()
On Error Resume Next
    Worksheets("Report").JournalEntryReport Co:=Me.Range("hdrCo").Value, Sys:=Me.Range("hdrSys").Value, JeType:=Me.Range("hdrJeType").Value, CtrlGrp:=Me.Range("hdrCtrlGrp").Value, FY:=Year(Me.Range("hdrPostDate").Value), Pd:=Month(Me.Range("hdrPostDate").Value)
    FixObjects
End Sub
Private Sub JournalsUpload_Click()
On Error GoTo error_handler
    If Not CheckUserAttributes() Then Login
    UploadJournalHeader
    UploadJournalDetails
    JournalEntryReport_Click
    FixObjects
    Exit Sub
error_handler:
    MsgBox ("Upload Error" & vbCrLf & Err.Number & ":" & Err.description)
End Sub
Private Sub FixObjects() ' See Utilities/fix_ActiveX_objects.bas
    ' JournalEntryReport
    Me.JournalEntryReport.Left = 469.5
    Me.JournalEntryReport.Width = 60
    Me.JournalEntryReport.Height = 21.75
    Me.JournalEntryReport.Top = 10.5
    Me.Shapes("JournalEntryReport").ScaleHeight 1.25, msoFalse, msoScaleFromTopLeft
    Me.Shapes("JournalEntryReport").ScaleHeight 0.8, msoFalse, msoScaleFromTopLeft

    ' JournalsUpload
    Me.JournalsUpload.Left = 387.75
    Me.JournalsUpload.Width = 60
    Me.JournalsUpload.Height = 21.75
    Me.JournalsUpload.Top = 9.75
    Me.Shapes("JournalsUpload").ScaleHeight 1.25, msoFalse, msoScaleFromTopLeft
    Me.Shapes("JournalsUpload").ScaleHeight 0.8, msoFalse, msoScaleFromTopLeft
End Sub
Private Sub UploadJournalHeader()
    Const search_string As String = "//text()"      ' XPath for all text elements in the document
    Dim response_element As MSXML2.IXMLDOMNode
    Dim request_parameters As String         ' String for building POST data
    Dim request_response As String
    Dim status As Integer   ' Status number return value
    Dim message_code As Integer     ' Message number return value

    status = 0
    request_parameters = "_PDL=" & g_sProductLine ' Start building POST data string with Product Line
    Select Case Me.Range("hdrFC").Value
        Case "A"        ' Add - ensure there's no JE #
            If Me.Range("hdrCtrlGrp").Value = "" Then
                request_parameters = request_parameters & "&_TKN=GL40.2&_EVT=ADD&_RTN=DATA&_TDS=IGNORE&FC=Add"
                Me.Range("hdrResponse").Value = ""  ' Clear the system response cell for the Header upload row on the spreadsheet
            Else
                Me.Range("hdrResponse").Value = "To add new, JE# must be blank."
                Exit Sub
            End If
        Case "C"        ' Change - ensure there is a JE #
            If Me.Range("hdrCtrlGrp").Value <> "" Then
                request_parameters = request_parameters & "&_TKN=GL40.2&_EVT=CHG&_RTN=DATA&_TDS=IGNORE&FC=Change"
                Me.Range("hdrResponse").Value = ""  ' Clear the system response cell for the Header upload row on the spreadsheet
            Else
                Me.Range("hdrResponse").Value = "To change JE header, must specify JE#."
                Exit Sub
            End If
        Case "D"        ' Delete - ensure there is a JE #
            If Me.Range("hdrCtrlGrp").Value <> "" Then
                ' deleting JE header requires hidden key (HK) value with Company, FY, Period, System, JE Type, JE #, and JE seq # formatted as 24 characters: ccccyyyymmsstccccccccqq
                request_parameters = request_parameters & "&_TKN=GL40.2&_EVT=CHG&_RTN=DATA&_TDS=IGNORE&FC=Delete&HK=" & _
                    format(Me.Range("hdrCo").Value, "0000") & format(Me.Range("hdrPostDate").Value, "yyyymm") & Me.Range("hdrSys").Value & _
                    Me.Range("hdrJeType").Value & format(Me.Range("hdrCtrlGrp").Value, "00000000") & format(Me.Range("hdrJeSeq").Value, "00")
                Me.Range("hdrResponse").Value = ""  ' Clear the system response cell for the Header upload row on the spreadsheet
            Else
                Me.Range("hdrResponse").Value = "To delete JE header, must specify JE#."
                Exit Sub
            End If
        Case ""         ' blank - skip header, continue to detail (add or change)
            Exit Sub
        Case Else       ' not 'A' or 'C' or blank??
            Me.Range("hdrResponse").Value = "Unknown function code - 'A', 'C' or 'D' only, blank to skip."
            Exit Sub
    End Select
    request_parameters = request_parameters & "&_f17=" & Me.Range("hdrCo").Value & "&_f20=" & _
        format(Me.Range("hdrPostDate").Value, "yyyy") & "&_f21=" & format(Me.Range("hdrPostDate").Value, "m") & _
        "&_f22=" & Me.Range("hdrSys").Value & "&_f24=" & Me.Range("hdrJeType").Value
    If Me.Range("hdrCtrlGrp").Value <> "" Then request_parameters = request_parameters & "&_f25=" & Me.Range("hdrCtrlGrp").Value ' Control group, if specified (will error if FC=A)
    If Me.Range("hdrJeSeq").Value <> "" Then request_parameters = request_parameters & "&_f26=" & Me.Range("hdrJeSeq").Value ' JE Sequence, if specified
    request_parameters = request_parameters & "&_f27=" & FilterForWeb(Left(Me.Range("hdrDesc").Value, 30)) ' JE Description (reqired) - Truncate to 30 characters
    If Me.Range("hdrSrc").Value <> "" Then request_parameters = request_parameters & "&_f30=" & Me.Range("hdrSrc").Value ' Source (optional, defaults to "JE")
    If Me.Range("hdrRef").Value <> "" Then request_parameters = request_parameters & "&_f34=" & FilterForWeb(Me.Range("hdrRef").Value) ' Reference (optional)
    If Me.Range("hdrAuRev").Value <> "" Then request_parameters = request_parameters & "&_f37=" & Me.Range("hdrAuRev").Value ' Auto Reverse (optional, defaults to "N")
    If Me.Range("hdrRevPd").Value <> "" Then request_parameters = request_parameters & "&_f38=" & Me.Range("hdrRevPd").Value ' Auto Reverse Period (optional, defaults to 0-next pd)
    If Me.Range("hdrDoc").Value <> "" Then request_parameters = request_parameters & "&_f42=" & FilterForWeb(Me.Range("hdrDoc").Value) ' Document (optional)
    request_parameters = request_parameters & "&_f48=" & format(Me.Range("hdrPostDate").Value, "yyyymmdd") ' Posting Date (required)
    If Me.Range("hdrTranDate").Value <> "" Then request_parameters = request_parameters & "&_f49=" & format(Me.Range("hdrTranDate").Value, "yyyymmdd") ' Transaction Date (optional)
    request_parameters = request_parameters & "&_OUT=XML&_EOT=TRUE" ' Send response as XML; (EOT=TRUE : ???)

    SetXMLObject ' Load page document into XML document object
    request_response = SendURL(request_parameters, "T")
    If Not g_oDom.LoadXML(request_response) Then
        If Me.Range("hdrFC").Value = "A" Then
            Me.Range("hdrFC").Value = "C"
            Me.Range("hdrResponse").Value = "Loading error - check if JE header exists before adding again."
        Else
            Me.Range("hdrResponse").Value = "Loading error - check JE report to confirm change."
        End If
        Exit Sub
    End If

    For Each response_element In g_oDom.SelectNodes(search_string) ' Decide what to do with each text element based on the parent node's name
        Select Case response_element.ParentNode.nodeName
            Case "Message"      ' This is the node name that returns the upload result
                Me.Range("hdrResponse").Value = Me.Range("hdrResponse").Value & response_element.Text
            Case "FldNbr"       ' This is the field number related to the response message
                Me.Range("hdrResponse").Value = Me.Range("hdrResponse").Value & "(" & response_element.Text & ")"
            Case "MsgNbr"       ' This return value refers to the message response - message number "000" is success
                message_code = Val(response_element.Text)
            Case "StatusNbr"    ' This is the status of the header upload.
                status = Val(response_element.Text)
            Case "_f25"         ' This is the Control Group (JE #)
                Me.Range("hdrCtrlGrp").Value = response_element.Text
            Case "_f26"         ' This is the JE Sequence Num
                Me.Range("hdrJeSeq").Value = response_element.Text
        End Select
    Next response_element

    If status = 1 And message_code = 0 Then ' If add/change/delete successful, delete the Function Code
        If Me.Range("hdrFC").Value = "D" Then Me.Range("hdrCtrlGrp").Value = "deleted (" & Me.Range("hdrCtrlGrp").Value & ")" ' mark if deleted
        Me.Range("hdrFC").Value = ""
    End If
End Sub
Private Sub UploadJournalDetails()
    Dim row As Long
    Dim response_element As MSXML2.IXMLDOMNode
    Dim request_parameters As String            ' String for building POST data
    Dim request_response As String
    Dim status As Integer                       ' Status number return value
    Dim message_code As Integer                 ' Message number return value
    Const search_string As String = "//text()"  ' XPath for all text elements in the document
    Const colFC As Integer = 1                  ' Columns for JE line data
    Const colToCo As Integer = 2
    Const colLine As Integer = 3
    Const colAcUnit As Integer = 4
    Const colAcct As Integer = 5
    Const colSubAcct As Integer = 6
    Const colActivity As Integer = 7
    Const colAcctCat As Integer = 8
    Const colAutoRev As Integer = 9
    Const colAmount As Integer = 10
    Const colDescription As Integer = 11
    Const colReference As Integer = 12
    Const colResponse As Integer = 13

    For row = 14 To Me.UsedRange.Rows.Count ' check for data to upload from row 14 through end of worksheet
        status = 0 ' Status 0 = continue on current row
        request_parameters = "_PDL=" & g_sProductLine ' Start building POST data string with Product Line
        Select Case Me.Cells(row, colFC).Value ' Decide how to treat line based on Function Code
            Case "A"        ' Add - ensure there's no Line #
                If Me.Cells(row, colLine).Value = "" Then
                    request_parameters = request_parameters & "&_TKN=GL40.1&_EVT=CHG&_RTN=DATA&_TDS=IGNORE&FC=Change"
                    Me.Cells(row, colResponse).Value = "" ' Clear the system response cell for the upload row
                Else
                    Me.Cells(row, colResponse).Value = "To add new, Line # must be blank."
                    status = 1 ' Status 1 = exit loop
                End If
            Case "C", "D"  ' Change or Delete - ensure there is a Line #
                If Me.Cells(row, colLine).Value <> "" Then
                    request_parameters = request_parameters & "&_TKN=GL40.1&_EVT=CHG&_RTN=DATA&_TDS=IGNORE&FC=Change"
                    Me.Cells(row, colResponse).Value = "" ' Clear the system response cell for the upload row
                Else
                    Me.Cells(row, colResponse).Value = "To change or delete JE line, must specify Line #."
                    status = 1 ' Status 1 = exit loop
                End If
            Case ""         ' blank - skip line
                status = 1 ' Status 1 = exit loop
            Case Else       ' not 'A' or 'C' or blank??
                Me.Cells(row, colResponse).Value = "Unknown function code - 'A', 'C' or 'D' only, blank to skip."
                status = 1 ' Status 1 = exit loop
        End Select
        If status = 0 Then ' if we haven't set status = 1 then continue with row upload
            request_parameters = request_parameters & "&_f39=" & Me.Range("hdrCo").Value  ' Company from Header
            request_parameters = request_parameters & "&_f44=" & format(Me.Range("hdrPostDate").Value, "yyyy")  ' Fiscal Year is year of Posted Date
            request_parameters = request_parameters & "&_f45=" & format(Me.Range("hdrPostDate").Value, "m")     ' Period is month of Posted Date
            request_parameters = request_parameters & "&_f46=" & Me.Range("hdrSys").Value ' System from Header
            request_parameters = request_parameters & "&_f48=" & Me.Range("hdrJeType").Value ' JE Type from Header
            request_parameters = request_parameters & "&_f49=" & Me.Range("hdrCtrlGrp").Value ' Control Group (JE#) from Header
            If Me.Range("hdrJeSeq").Value <> "" Then request_parameters = request_parameters & "&_f50=" & Me.Range("hdrJeSeq").Value ' JE Sequence # from Header, if there is one
            request_parameters = request_parameters & "&_f67r0=" & Me.Cells(row, colFC).Value ' Detail line Function Code
            If Me.Cells(row, colLine).Value <> "" Then request_parameters = request_parameters & "&_f79r0=" & Me.Cells(row, colLine).Value ' JE line number, if there is one
            If Me.Cells(row, colToCo).Value = "" Then request_parameters = request_parameters & "&_f68r0=" & Me.Range("hdrCo").Value Else request_parameters = request_parameters & "&_f68r0=" & Me.Cells(row, colToCo).Value ' To Company, if specified; else Company from Header
            request_parameters = request_parameters & "&_f69r0=" & Me.Cells(row, colAcUnit).Value ' Accounting Unit (cost center)
            request_parameters = request_parameters & "&_f70r0=" & Me.Cells(row, colAcct).Value ' GL Account
            If Me.Cells(row, colSubAcct).Value <> "" Then request_parameters = request_parameters & "&_f71r0=" & Me.Cells(row, colSubAcct).Value ' GL Sub Account, if specified
            If Me.Cells(row, colActivity).Value <> "" Then request_parameters = request_parameters & "&_f73r0=" & Me.Cells(row, colActivity).Value ' Activity code, if specified
            If Me.Cells(row, colAcctCat).Value <> "" Then request_parameters = request_parameters & "&_f74r0=" & Me.Cells(row, colAcctCat).Value ' Account Category code, if specified
            If Me.Cells(row, colAutoRev).Value <> "" Then request_parameters = request_parameters & "&_f86r0=" & Me.Cells(row, colAutoRev).Value Else request_parameters = request_parameters & "&_f86r0=" & Me.Range("hdrAuRev").Value ' Auto Reverse flage, if specified; else Auto Rev from Header
            request_parameters = request_parameters & "&_f75r0=" & Me.Cells(row, colAmount).Value ' Transaction Amount
            request_parameters = request_parameters & "&_f81r0=" & FilterForWeb(Left(Me.Cells(row, colDescription).Value, 30)) ' JE Description - Truncate to 30 characters
            If Me.Range("hdrSrc").Value <> "" Then request_parameters = request_parameters & "&_f89r0=" & Me.Range("hdrSrc").Value ' Source from Header, if specified
            If Me.Cells(row, colReference).Value <> "" Then request_parameters = request_parameters & "&_f88r0=" & FilterForWeb(Me.Cells(row, colReference).Value) ' Reference field, if specified
            request_parameters = request_parameters & "&_OUT=XML&_EOT=TRUE&_INITDTL=TRUE" ' Send response in XML; (EOT=TRUE : ???); bypass requiring an inquire before change

            SetXMLObject
            request_response = SendURL(request_parameters, "T")
            If Not g_oDom.LoadXML(request_response) Then
                If Me.Cells(row, colFC).Value = "A" Then
                    Me.Cells(row, colFC).Value = "C"
                    Me.Cells(row, colResponse).Value = "Loading error - check if line exists before adding again."
                Else
                    Me.Cells(row, colResponse).Value = "Loading error - check JE report to confirm change."
                End If
                Exit Sub
            End If

            For Each response_element In g_oDom.SelectNodes(search_string) ' Decide what to do with each text element based on the parent node's name
                Select Case response_element.ParentNode.nodeName
                    Case "Message"      ' This is the node name that returns the upload result
                        Me.Cells(row, colResponse).Value = Me.Cells(row, colResponse).Value & response_element.Text
                    Case "FldNbr"       ' This is the field number related to the response message
                        Me.Cells(row, colResponse).Value = Me.Cells(row, colResponse).Value & "(" & response_element.Text & ")"
                    Case "MsgNbr"       ' This return value refers to the message response - message number "000" is success
                        message_code = Val(response_element.Text)
                    Case "StatusNbr"    ' This is the status of the header upload. "001" = received
                        status = Val(response_element.Text)
                    Case "_f79r0"         ' This is the line number
                        Me.Cells(row, colLine).Value = response_element.Text
                End Select
            Next response_element

            If status = 1 And message_code = 0 Then ' note that status has been repurposed from server response ' If add/change/delete successful, delete the Function Code
                If Me.Cells(row, colFC).Value = "D" Then Me.Cells(row, colLine).Value = "deleted (" & Me.Cells(row, colLine).Value & ")" ' and indicate successful delete
                Me.Cells(row, colFC).Value = ""
            End If
        End If
    Next row
End Sub
