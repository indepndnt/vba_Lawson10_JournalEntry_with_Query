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
Private Type ProfileAttributes
    lawsonusername As String
    id As String
    lawsonuserlogin As String
End Type
Private Type DebitCredit
    debit As Currency
    credit As Currency
End Type
Public Sub inJournalEditRpt() ' Used to run only inGL240 - point 'Report' button on Upload tab to here.
On Error GoTo errHandler
    If Not CheckUserAttributes() Then Login
    If Not inGL240() Then ' Run GL240 report
        MsgBox "Report Query Error"
    End If
Exit Sub
errHandler:
    MsgBox ("Report Error" & vbCrLf & Err.Number & ":" & Err.Description)
End Sub
Public Function inGL240() As Boolean

    Dim s As String  ' String to build POST data
    Dim row As Long  ' Row # for report worksheet output
    Dim line As Long ' Row # for GL Detail output
    Dim lLines As Long ' Total detail lines in JE
    Dim sOutOperator As String ' JE Operator name
    Dim sOutYesNo As String ' Yes/No as string
    Dim sOutAutoRevPd As String ' Auto reverse period as string
    Dim sJeType As String ' JE Type from JE Header
    Dim sCtrlGrp As String ' JE Control Group from JE Header
    Dim sJeSeq As String ' JE Sequence from JE Header
    Dim lAmount As DebitCredit
    Dim lTotal As DebitCredit
    Dim lReverse As DebitCredit
    Dim lBase As DebitCredit
    Dim lUnit As DebitCredit
    Dim cAmount As DebitCredit

    inGL240 = False ' Defaults if no result
    If Sheet3.Range("hdrCtrlGrp").Value = "" Then
        MsgBox ("Cannot report without JE number")
        Exit Function
    End If
    If Sheet3.Range("hdrPostDate").Value = "" Then
        MsgBox ("Cannot report without Posting Date")
        Exit Function
    End If
    If Sheet3.Range("hdrCo").Value = "" Then
        MsgBox ("Cannot report without Company value")
        Exit Function
    End If
    If Sheet3.Range("hdrJeType").Value = "" Then
        MsgBox ("Cannot report without JE Type")
        Exit Function
    End If

    Dim prof As ProfileAttributes
    prof = fProfileUser()

    s = "PROD=" & g_sProductLine ' Start building POST data string with Product Line
    s = s & "&FILE=GLCONTROL&INDEX=GLCSET1" ' Table GLCONTROL, criteria set GLCSET1: key = co=fy=pd=sys=jetype=ctrlgrp
    s = s & "&KEY=" & FilterForWeb(Sheet3.Range("hdrCo").Value & "=" & format(Sheet3.Range("hdrPostDate"), "yyyy=mm=") & Sheet3.Range("hdrSys").Value & "=" & Sheet3.Range("hdrJeType").Value & "=" & Sheet3.Range("hdrCtrlGrp").Value)
    s = s & "&FIELD=" & FilterForWeb("COMPANY;COMPANY.NAME;COMPANY.CURRENCY-CODE;FISCAL-YEAR;ACCT-PERIOD;SYSTEM;JE-TYPE;CONTROL-GROUP;JE-SEQUENCE;DESCRIPTION;STATUS,xlt;HOLD-CODE;HOLD-REM-OPER;OPERATOR;POSTING-DATE;DATE;AUTO-REV;AUTO-REV-PD;REFERENCE;DOCUMENT-NBR;JRNL-BOOK-NBR;NBR-LINES")
    s = s & "&OUT=XML&NEXT=FALSE" ' repl"&OUT=CSV&DELIM=%09&NOHEADER=TRUE" NEXT=FALSE means don't give me the RELOAD string.
    s = s & "&MAX=1&keyUsage=PARAM" ' Only give me 1 record.

    s = SendURL(s, "D")
    SetXMLObject ' Load IE page document into XML document object
    If Not g_oDom.LoadXML(s) Then
        inGL240 = False ' If we couldn't load g_oDom with the Lawson output then exit with an error - there's no data.
        Exit Function
    End If
    If Not inXmlDme Then ' do we have a /DME xml document?
        Sheet3.Range("hdrResponse").Value = inXmlData("/ERROR/MSG", 1) ' Error message from GLCONTROL query
        inGL240 = True
        Exit Function
    End If

' Header query columns:
'  1:COMPANY       2:COMPANY.NAME  3:COMPANY.CURRENCY-CODE  4:FISCAL-YEAR     5:ACCT-PERIOD  6:SYSTEM         7:JE-TYPE  8:CONTROL-GROUP  9:JE-SEQUENCE
' 10:DESCRIPTION  11:STATUS,xlt   12:HOLD-CODE             13:HOLD-REM-OPER  14:OPERATOR    15:POSTING-DATE  16:DATE    17:AUTO-REV      18:AUTO-REV-PD
' 19:REFERENCE    20:DOCUMENT-NBR 21:JRNL-BOOK-NBR         22:NBR-LINES
    lLines = Int("0" & Trim(inQueryCell(1, 22))) ' Number of lines as reported by GLCONTROL query
    row = 12    ' Header goes through row 11, rows 12-end vary
    Call fDeleteFrom(Sheet2.Cells(row, 1)) ' delete any previous outputs and reset UsedRange

    sJeType = inQueryCell(1, 7)
    sCtrlGrp = inQueryCell(1, 8)
    sJeSeq = inQueryCell(1, 9)
    sOutOperator = inQueryCell(1, 14)
    If sOutOperator = prof.lawsonuserlogin Then
        sOutOperator = prof.lawsonusername
    End If
    If inQueryCell(1, 17) = "Y" Then sOutYesNo = "Yes" Else sOutYesNo = "No "
    sOutAutoRevPd = inQueryCell(1, 18)
    Sheet2.Cells(1, 1).Value = "GL240 Date " & format(Now(), "MM/DD/YY") & Space(26) & "Company " & Right(Space(4) & inQueryCell(1, 1), 4) & _
        " - " & Left(inQueryCell(1, 2) & Space(32), 32) & Left(inQueryCell(1, 3) & Space(27), 27)
    Sheet2.Cells(2, 1).Value = "      Time " & format(Now(), "HH:MM") & Space(29) & "Journal Edit Listing"
    Sheet2.Cells(3, 1).Value = Space(45) & "For Fiscal Year " & inQueryCell(1, 4) & " - Periods " & format(inQueryCell(1, 5), "00") & _
        " - " & format(inQueryCell(1, 5), "00")
    Sheet2.Cells(4, 1).Value = ""
    Sheet2.Cells(5, 1).Value = " Journal            " & inQueryCell(1, 6) & " " & sJeType & _
        Right("         " & sCtrlGrp, 9) & "-" & format(sJeSeq, "00") & " " & Left(inQueryCell(1, 10) & Space(36), 36) & _
        "Fiscal Year    " & inQueryCell(1, 4) & "          Period     " & inQueryCell(1, 5)
    Sheet2.Cells(6, 1).Value = "   Status           " & Left(inQueryCell(1, 11) & Space(17), 17) & "Hold Code " & _
        Left(inQueryCell(1, 12) & Space(10), 10) & "Hold Removal Operator " & Left(inQueryCell(1, 13) & Space(13), 13) & _
        "Operator       " & sOutOperator
    Sheet2.Cells(7, 1).Value = "   Posting Date     " & format(inQueryCell(1, 15), "MM/DD/YY") & _
        "         Transaction Date    " & format(inQueryCell(1, 16), "MM/DD/YY") & "          Reverse  " & sOutYesNo & "     Reverse Pd     " & sOutAutoRevPd
    Sheet2.Cells(8, 1).Value = "   Reference        " & Left(inQueryCell(1, 19) & Space(16), 16) & " Document            " & _
        Left(inQueryCell(1, 20) & Space(34), 34) & " Journal Book   " & Left(inQueryCell(1, 21) & Space(9), 9)
    Sheet2.Cells(9, 1).Value = ""
    Sheet2.Cells(10, 1).Value = " Line   Co           Account                     Activity         Ref    SC Rvs                    Debit                    Credit"
    Sheet2.Cells(11, 1).Value = "------ ---- --------------------------- --------------------- ---------- -- --- ------------------------- -------------------------"

    s = "PROD=" & g_sProductLine ' Start building POST data string with Product Line
    s = s & "&FILE=GLTRANS&INDEX=GLTSET1" ' ' Table GLTRANS, criteria set GLTSET1: key = co=fy=pd=ctrlgrp=sys=jetype
    s = s & "&KEY=" & FilterForWeb(Sheet3.Range("hdrCo").Value & "=" & format(Sheet3.Range("hdrPostDate"), "yyyy=mm=") & Sheet3.Range("hdrCtrlGrp").Value & "=" & Sheet3.Range("hdrSys").Value & "=" & Sheet3.Range("hdrJeType").Value)
    s = s & "&FIELD=" & FilterForWeb("LINE-NBR;TO-COMPANY;ACCT-UNIT;ACCOUNT;SUB-ACCOUNT;ACTIVITY;ACCT-CATEGORY;REFERENCE;SOURCE-CODE;AUTO-REV;TRAN-AMOUNT;CHART-DETAIL.ACCOUNT-DESC;DESCRIPTION;BASE-AMOUNT;UNITS-AMOUNT")
    s = s & "&OUT=XML&NEXT=FALSE" ' NEXT=FALSE means don't give me the RELOAD string.
    s = s & "&MAX=" & line & "&keyUsage=PARAM" ' Give me as many records as header reported there are
    
    s = SendURL(s, "D")
    If Not g_oDom.LoadXML(s) Then
        inGL240 = False ' If we couldn't load g_oDom with the Lawson output then exit with an error - there's no data.
        Exit Function
    End If
    If Not inXmlDme Then ' do we have a /DME xml document?
        Sheet3.Range("hdrResponse").Value = inXmlData("/ERROR/MSG", 1) ' Error message from GLCONTROL query
        inGL240 = True
        Exit Function
    End If

    lBase = dcSet(0, 0)
    lReverse = dcSet(0, 0)
    lTotal = dcSet(0, 0)
    lUnit = dcSet(0, 0)

' Detail lines query columns:
'  1:LINE-NBR      2:TO-COMPANY   3:ACCT-UNIT    4:ACCOUNT       5:SUB-ACCOUNT   6:ACTIVITY     7:ACCT-CATEGORY  8:REFERENCE  9:SOURCE-CODE  10:AUTO-REV
' 11:TRAN-AMOUNT  12:CHART-DETAIL.ACCOUNT-DESC  13:DESCRIPTION  14:BASE-AMOUNT  15:UNITS-AMOUNT
    For line = 1 To lLines
        lAmount = inCurrency(inQueryCell(line, 11))
        lTotal = dcAdd(lTotal, lAmount) ' Add line amounts for total
        lBase = dcAdd(lBase, inCurrency(inQueryCell(line, 14))) ' Add Base Amounts for total
        lUnit = dcAdd(lUnit, inCurrency(inQueryCell(line, 15))) ' Add Unit Amounts for total

        If inQueryCell(line, 10) = "Y" Then  ' Set "Yes", "No" or blank string for Auto-Reverse line and add to total
            sOutYesNo = "Yes"
            lReverse = dcAdd(lReverse, lAmount)
        ElseIf inQueryCell(line, 10) = "N" Then
            sOutYesNo = "No "
        Else
            sOutYesNo = "   "
        End If
        
        If inQueryCell(line, 1) <> "" Then
            Sheet2.Cells(row, 1).Value = Right("      " & inQueryCell(line, 1), 6) & " " & Right("    " & inQueryCell(line, 2), 4) & " " & _
                Left(inQueryCell(line, 3) & Space(17), 17) & format(inQueryCell(line, 4), "00000") & "-" & format(inQueryCell(line, 5), "0000") & _
                " " & Left(inQueryCell(line, 6) & Space(16), 16) & Left(inQueryCell(line, 7) & Space(5), 5) & " " & Left(inQueryCell(line, 8) & _
                Space(10), 10) & " " & Left(inQueryCell(line, 9) & "  ", 2) & " " & sOutYesNo & Right(Space(26) & format(lAmount.debit, "#,##0.00;;\ "), 26) & _
                Right(Space(26) & format(lAmount.credit, "#,##0.00;;\ "), 26)
            row = row + 1
            Sheet2.Cells(row, 1).Value = Left(inQueryCell(line, 12) & Space(27), 27) & " " & inQueryCell(line, 13)
            row = row + 2
        End If
    Next line
    Sheet2.Cells(row, 1).Value = "*** Totals For Journal entry " & sJeType & "-" & Right("       " & sCtrlGrp, 8) & "-" & _
        format(sJeSeq, "00") & "                              Debits                   Credits                Difference"
    row = row + 1
    Sheet2.Cells(row, 1).Value = Space(37) & "Base . . . . . . ." & Right(Space(24) & format(lBase.debit, "#,##0.00"), 24) & _
        Right(Space(26) & format(lBase.credit, "#,##0.00"), 26) & Right(Space(26) & format(dcNet(lBase), "#,##0.00"), 26)
    row = row + 1
    Sheet2.Cells(row, 1).Value = Space(37) & "Reverse  . . . . ." & Right(Space(24) & format(lReverse.debit, "#,##0.00"), 24) & _
        Right(Space(26) & format(lReverse.credit, "#,##0.00"), 26) & Right(Space(26) & format(dcNet(lReverse), "#,##0.00"), 26)
    row = row + 1
    Sheet2.Cells(row, 1).Value = Space(37) & "Entered  . . . . ." & Right(Space(24) & format(lTotal.debit, "#,##0.00"), 24) & _
        Right(Space(26) & format(lTotal.credit, "#,##0.00"), 26) & Right(Space(26) & format(dcNet(lTotal), "#,##0.00"), 26)
    row = row + 1
    Sheet2.Cells(row, 1).Value = Space(37) & "Unit . . . . . . ." & Right(Space(24) & format(lUnit.debit, "#,##0.00"), 24) & _
        Right(Space(26) & format(lUnit.credit, "#,##0.00"), 26) & Right(Space(26) & format(dcNet(lUnit), "#,##0.00"), 26)

    inGL240 = True

End Function
Private Function inCurrency(ByVal s As String) As DebitCredit
    Dim bNegative As Boolean
    bNegative = s Like "*-"
    s = Replace(s, "-", "")
    s = Replace(s, " ", "")
    inCurrency.debit = 0
    inCurrency.credit = 0
    If bNegative Then
        inCurrency.credit = Val(s)
    Else
        inCurrency.debit = Val(s)
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
Private Function fProfileUser() As ProfileAttributes
    Dim s As String
    Dim xmlNode As MSXML2.IXMLDOMNode
    Dim XmlNodeList As MSXML2.IXMLDOMNodeList

    s = g_sServer & "/servlet/Profile?section=attributes"
    s = SendURL(s, "X")

    If Not g_oDom.LoadXML(s) Then
        fProfileUser.lawsonusername = ""
        Exit Function
    End If
    If Not inXmlDme("/PROFILE") Then
        fProfileUser.lawsonusername = ""
        Exit Function
    End If

    Set XmlNodeList = g_oDom.DocumentElement.SelectNodes("//ATTR")
    For Each xmlNode In XmlNodeList
        Select Case xmlNode.Attributes.getNamedItem("name").Text
            Case "lawsonusername"
                fProfileUser.lawsonusername = xmlNode.Attributes.getNamedItem("value").Text
            Case "id"
                fProfileUser.id = xmlNode.Attributes.getNamedItem("value").Text
            Case "lawsonuserlogin"
                fProfileUser.lawsonuserlogin = xmlNode.Attributes.getNamedItem("value").Text
        End Select
    Next xmlNode

End Function
