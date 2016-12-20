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
Private Type ProfileAttributes
    lawsonusername As String ' dddddd\uuuuuuuu
    id As String ' uuuuuuuu
    lawsonuserlogin As String ' NT00000nnn
End Type
Private Type DebitCredit
    debit As Currency
    credit As Currency
End Type
Public Sub inGL240(ByVal Co As Long, ByVal Sys As String, ByVal JeType As String, ByVal CtrlGrp As Long, ByVal FY As Long, ByVal Pd As Integer)

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
    Dim prof As ProfileAttributes

    If Not CheckUserAttributes() Then Login
    prof = fProfileUser()

    s = "PROD=" & g_sProductLine & "&FILE=GLCONTROL&INDEX=GLCSET1"
    s = s & "&KEY=" & FilterForWeb(Co & "=" & FY & "=" & Pd & "=" & Sys & "=" & JeType & "=" & CtrlGrp)
    s = s & "&FIELD=" & FilterForWeb("COMPANY;COMPANY.NAME;COMPANY.CURRENCY-CODE;FISCAL-YEAR;ACCT-PERIOD;SYSTEM;JE-TYPE;CONTROL-GROUP;JE-SEQUENCE;DESCRIPTION;STATUS,xlt;HOLD-CODE;HOLD-REM-OPER;OPERATOR;POSTING-DATE;DATE;AUTO-REV;AUTO-REV-PD;REFERENCE;DOCUMENT-NBR;JRNL-BOOK-NBR;NBR-LINES")
    s = s & "&OUT=XML&NEXT=FALSE&MAX=1&keyUsage=PARAM"

    s = SendURL(s, "D")
    SetXMLObject ' Load IE page document into XML document object
    If Not g_oDom.LoadXML(s) Then Exit Sub
    If Not inXmlDme Then ' do we have a /DME xml document?
        Me.Range("B5").Value = inXmlData("/ERROR/MSG", 1) ' Error message from GLCONTROL query
        Exit Sub
    End If

' Header query columns:
'  1:COMPANY       2:COMPANY.NAME  3:COMPANY.CURRENCY-CODE  4:FISCAL-YEAR     5:ACCT-PERIOD  6:SYSTEM         7:JE-TYPE  8:CONTROL-GROUP  9:JE-SEQUENCE
' 10:DESCRIPTION  11:STATUS,xlt   12:HOLD-CODE             13:HOLD-REM-OPER  14:OPERATOR    15:POSTING-DATE  16:DATE    17:AUTO-REV      18:AUTO-REV-PD
' 19:REFERENCE    20:DOCUMENT-NBR 21:JRNL-BOOK-NBR         22:NBR-LINES
    Dim rArray() As String
    Dim qArray() As String
    Call inQueryArray(qArray)

    lLines = Int("0" & Trim(qArray(0, 21))) ' Number of lines as reported by GLCONTROL query
    ReDim rArray(lLines * 3 + 16)
    row = 11    ' Header goes through row 11, rows 12-end vary
    Call fDeleteFrom(Me.Cells(row + 1, 1)) ' delete any previous outputs and reset UsedRange

    sJeType = qArray(0, 6)
    sCtrlGrp = qArray(0, 7)
    sJeSeq = qArray(0, 8)
    sOutOperator = qArray(0, 13)
    If sOutOperator = prof.lawsonuserlogin Then
        sOutOperator = prof.lawsonusername
    End If
    If qArray(0, 16) = "Y" Then sOutYesNo = "Yes" Else sOutYesNo = "No "
    sOutAutoRevPd = qArray(0, 17)
    rArray(0) = "GL240 Date " & format(Now(), "MM/DD/YY") & Space(26) & "Company " & Right(Space(4) & qArray(0, 0), 4) & _
        " - " & Left(qArray(0, 1) & Space(32), 32) & Left(qArray(0, 2) & Space(27), 27)
    rArray(1) = "      Time " & format(Now(), "HH:MM") & Space(29) & "Journal Edit Listing"
    rArray(2) = Space(45) & "For Fiscal Year " & qArray(0, 3) & " - Periods " & format(qArray(0, 4), "00") & _
        " - " & format(qArray(0, 4), "00")
    rArray(3) = ""
    rArray(4) = " Journal            " & qArray(0, 5) & " " & sJeType & _
        Right("         " & sCtrlGrp, 9) & "-" & format(sJeSeq, "00") & " " & Left(qArray(0, 9) & Space(36), 36) & _
        "Fiscal Year    " & qArray(0, 3) & "          Period     " & qArray(0, 4)
    rArray(5) = "   Status           " & Left(qArray(0, 10) & Space(17), 17) & "Hold Code " & _
        Left(qArray(0, 11) & Space(10), 10) & "Hold Removal Operator " & Left(qArray(0, 12) & Space(13), 13) & _
        "Operator       " & sOutOperator
    rArray(6) = "   Posting Date     " & format(qArray(0, 14), "MM/DD/YY") & _
        "         Transaction Date    " & format(qArray(0, 15), "MM/DD/YY") & "          Reverse  " & sOutYesNo & "     Reverse Pd     " & sOutAutoRevPd
    rArray(7) = "   Reference        " & Left(qArray(0, 18) & Space(16), 16) & " Document            " & _
        Left(qArray(0, 19) & Space(34), 34) & " Journal Book   " & Left(qArray(0, 20) & Space(9), 9)
    rArray(8) = ""
    rArray(9) = " Line   Co           Account                     Activity         Ref    SC Rvs                    Debit                    Credit"
    rArray(10) = "------ ---- --------------------------- --------------------- ---------- -- --- ------------------------- -------------------------"

    s = "PROD=" & g_sProductLine ' Start building POST data string with Product Line
    s = s & "&FILE=GLTRANS&INDEX=GLTSET1" ' ' Table GLTRANS, criteria set GLTSET1: key = co=fy=pd=ctrlgrp=sys=jetype
    s = s & "&KEY=" & FilterForWeb(Co & "=" & FY & "=" & Pd & "=" & CtrlGrp & "=" & Sys & "=" & JeType)
    s = s & "&FIELD=" & FilterForWeb("LINE-NBR;TO-COMPANY;ACCT-UNIT;ACCOUNT;SUB-ACCOUNT;ACTIVITY;ACCT-CATEGORY;REFERENCE;SOURCE-CODE;AUTO-REV;TRAN-AMOUNT;CHART-DETAIL.ACCOUNT-DESC;DESCRIPTION;BASE-AMOUNT;UNITS-AMOUNT")
    s = s & "&OUT=XML&NEXT=FALSE&MAX=10000&keyUsage=PARAM"
    
    s = SendURL(s, "D")
    If Not g_oDom.LoadXML(s) Then Exit Sub
    If Not inXmlDme Then ' do we have a /DME xml document?
        Me.Range("B6").Value = inXmlData("/ERROR/MSG", 1) ' Error message from GLCONTROL query
        Exit Sub
    End If

    lBase = dcSet(0, 0)
    lReverse = dcSet(0, 0)
    lTotal = dcSet(0, 0)
    lUnit = dcSet(0, 0)
    Call inQueryArray(qArray)

' Detail lines query columns:
'  1:LINE-NBR      2:TO-COMPANY   3:ACCT-UNIT    4:ACCOUNT       5:SUB-ACCOUNT   6:ACTIVITY     7:ACCT-CATEGORY  8:REFERENCE  9:SOURCE-CODE  10:AUTO-REV
' 11:TRAN-AMOUNT  12:CHART-DETAIL.ACCOUNT-DESC  13:DESCRIPTION  14:BASE-AMOUNT  15:UNITS-AMOUNT
    For line = 0 To lLines - 1
        lAmount = inCurrency(qArray(line, 10))
        lTotal = dcAdd(lTotal, lAmount) ' Add line amounts for total
        lBase = dcAdd(lBase, inCurrency(qArray(line, 13))) ' Add Base Amounts for total
        lUnit = dcAdd(lUnit, inCurrency(qArray(line, 14))) ' Add Unit Amounts for total

        If qArray(line, 9) = "Y" Then  ' Set "Yes", "No" or blank string for Auto-Reverse line and add to total
            sOutYesNo = "Yes"
            lReverse = dcAdd(lReverse, lAmount)
        ElseIf qArray(line, 9) = "N" Then
            sOutYesNo = "No "
        Else
            sOutYesNo = "   "
        End If
        
        If qArray(line, 0) <> "" Then
            rArray(row) = Right("      " & qArray(line, 0), 6) & " " & Right("    " & qArray(line, 1), 4) & " " & _
                Left(qArray(line, 2) & Space(17), 17) & format(qArray(line, 3), "00000") & "-" & format(qArray(line, 4), "0000") & _
                " " & Left(qArray(line, 5) & Space(16), 16) & Left(qArray(line, 6) & Space(5), 5) & " " & Left(qArray(line, 7) & _
                Space(10), 10) & " " & Left(qArray(line, 8) & "  ", 2) & " " & sOutYesNo & Right(Space(26) & format(lAmount.debit, "#,##0.00;;\ "), 26) & _
                Right(Space(26) & format(lAmount.credit, "#,##0.00;;\ "), 26)
            rArray(row + 1) = Left(qArray(line, 11) & Space(27), 27) & " " & qArray(line, 12)
            rArray(row + 2) = ""
            row = row + 3
        End If
    Next line
    rArray(row) = "*** Totals For Journal entry " & sJeType & "-" & Right("       " & sCtrlGrp, 8) & "-" & _
        format(sJeSeq, "00") & "                              Debits                   Credits                Difference"
    rArray(row + 1) = Space(37) & "Base . . . . . . ." & Right(Space(24) & format(lBase.debit, "#,##0.00"), 24) & _
        Right(Space(26) & format(lBase.credit, "#,##0.00"), 26) & Right(Space(26) & format(dcNet(lBase), "#,##0.00"), 26)
    rArray(row + 2) = Space(37) & "Reverse  . . . . ." & Right(Space(24) & format(lReverse.debit, "#,##0.00"), 24) & _
        Right(Space(26) & format(lReverse.credit, "#,##0.00"), 26) & Right(Space(26) & format(dcNet(lReverse), "#,##0.00"), 26)
    rArray(row + 3) = Space(37) & "Entered  . . . . ." & Right(Space(24) & format(lTotal.debit, "#,##0.00"), 24) & _
        Right(Space(26) & format(lTotal.credit, "#,##0.00"), 26) & Right(Space(26) & format(dcNet(lTotal), "#,##0.00"), 26)
    rArray(row + 4) = Space(37) & "Unit . . . . . . ." & Right(Space(24) & format(lUnit.debit, "#,##0.00"), 24) & _
        Right(Space(26) & format(lUnit.credit, "#,##0.00"), 26) & Right(Space(26) & format(dcNet(lUnit), "#,##0.00"), 26)
    row = row + 4
    Me.Range("A1").Resize(UBound(rArray)).Value = Application.Transpose(rArray)

End Sub
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
