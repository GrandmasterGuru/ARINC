Dim currdate As Date
Dim currfolder As String
Dim separator As String
Dim spacer As String
Dim seqin As Integer
Dim seqout As Integer
Dim newmonth As Boolean
Dim fso As New FileSystemObject, fldr As Folder, fl As File
Dim inbound()
Dim dbDataConn As New ADODB.Connection   'object declared and created.
Dim dbData As New ADODB.Recordset
Private Sub Form_Load()

    separator = ""
    For i = 1 To 7
        separator = separator & "##########"
    Next
    separator = vbCrLf & separator & vbCrLf
    spacer = ""
    For i = 1 To 13
        spacer = spacer & vbCrLf
    Next
    currfldrname = DatePart("yyyy", Now) & "_"   'e.g., 2009_
    x = DatePart("m", Now)
    If Len(x) = 1 Then x = "0" & x
    currmnth = x
    currfldrname = currfldrname & x & "_" 'e.g., 2009_11_
    x = DatePart("d", Now)
    If Len(x) = 1 Then x = "0" & x
    currfldrname = currfldrname & x   'e.g., 2009_11_21
    If fso.FolderExists("C:\ARINC\Archives\In\" & currfldrname) Then
        hiseq = 0
        Set fldrIn = fso.GetFolder("C:\ARINC\Archives\In\" & currfldrname)
        For Each Fil In fldrIn.Files
            seqin = Left(Fil.Name, InStr(1, Fil.Name, "_") - 1)
            If seqin > hiseq Then hiseq = seqin
        Next
        seqin = hiseq
        hiseq = 0
        Set fldrOut = fso.GetFolder("C:\ARINC\Archives\Out\" & currfldrname)
        For Each Fil In fldrIn.Files
            seqout = Trim(Left(Fil.Name, InStr(1, Fil.Name, "_") - 1))
            If seqout > hiseq Then hiseq = seqin
        Next
        seqout = hiseq
    Else
        fso.CreateFolder ("C:\ARINC\Archives\In\" & currfldrname)
        fso.CreateFolder ("C:\ARINC\Archives\Out\" & currfldrname)
        seqin = 0
        seqout = 0
    End If
    newmonth = False
    currdate = FormatDateTime(Now, vbShortDate)
    dbDataConn.ConnectionString = "DSN=mysql; userid=pegasusanc; password="
    dbDataConn.Open
    
End Sub

Private Sub Pegasus_WillMove(ByVal adReason As ADODB.EventReasonEnum, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)

End Sub

Private Sub Timer1_Timer()
    Dim fso As New FileSystemObject, fldr As Folder, fl As File
   ' First, check to see if we have passed midnight
    newfldrname = DatePart("yyyy", Now) & "_"   '2009_
    x = DatePart("m", Now)
    If Len(x) = 1 Then x = "0" & x
    currmnth = x
    newfldrname = newfldrname & x & "_" '2009_11_
    today = DatePart("d", Now)
    If today = 1 Then
        If newmonth = False Then
            newmonth = True
            seqin = 0
            seqout = 0
            fourmonthsago = DateAdd("m", -4, Now)
            newfldrname = DatePart("yyyy", fourmonthsago) & "_"   '2009_
            x = DatePart("m", fourmonthsago)
            If Len(x) = 1 Then x = "0" & x
            currmnth = x
            newfldrname = newfldrname & x & "_" & "01"  '2009_11_
            For i = 1 To 31
                If fso.FolderExists("C:\ARINC\Archives\In\" & newfldrname) Then fso.DeleteFolder ("C:\ARINC\Archives\In\" & newfldrname)
                If fso.FolderExists("C:\ARINC\Archives\Out\" & newfldrname) Then fso.DeleteFolder ("C:\ARINC\Archives\Out\" & newfldrname)
                newfldrname = Left(newfldrname, 8)  ' 2009_08_
                x = i + 1
                If Len(x) = 1 Then x = "0" & x
                newfldrname = newfldrname & x   '2009_08_02
            Next
        End If
    Else
        newmonth = False
    End If
    If Len(today) = 1 Then today = "0" & today
    newfldrname = newfldrname & today   '2009_11_21
    If DateDiff("d", currdate, FormatDateTime(Now, vbShortDate)) <> 0 Or Not fso.FolderExists("C:\ARINC\Archives\In\" & newfldrname) Then
        ' We have gone past midnight
        ' Need to create new inbound and outbound folders
        ' Folder name format is yyyy_mm_dd
        fso.CreateFolder ("C:\ARINC\Archives\In\" & newfldrname)
        fso.CreateFolder ("C:\ARINC\Archives\Out\" & newfldrname)
        currdate = FormatDateTime(Now, vbShortDate)
    End If
    currfolder = newfldrname
    found = False
    Set fldrIn = fso.GetFolder("c:\ARINC\Incoming")
    On Error Resume Next
    For Each Fil In fldrIn.Files
        found = True
        seqin = seqin + 1
        headline = "Rcv: " & Str(seqin) & " file: " & Fil.Name & " " & FormatDateTime(Now, vbShortDate) & " " & FormatDateTime(Now, vbLongTime) & vbCrLf
'        toline = msgin.ReadLine
'        fromline = msgin.ReadLine
'        If InStr(1, fromline, "QIFNMCA") > 0 Or InStr(1, fromline, "QIFPMCA") > 0 Or InStr(1, fromline, "PEKUDCA") > 0 Then
'            For Each pr In Printers
'                If pr.DeviceName = "CANON" Then Set Printer = pr
'            Next
'        Else
'        End If
'        Printer.Print fromline
'        Printer.Print toline
        ReDim inbound(0)
        keep = True
        foundaddr = False
        Open "\\127.0.0.1\ARINC_IN" For Output As #1
        Print #1, headline
        ReDim Preserve inbound(UBound(inbound) + 1)
        inbound(UBound(inbound)) = Fil.Name
        Set msgin = Fil.OpenAsTextStream(ForReading)
        While Not msgin.AtEndOfStream
            theline = msgin.ReadLine
            If Len(theline) < 3 Then theline = "  "
            Print #1, theline
            If keep = True Then
                ReDim Preserve inbound(UBound(inbound) + 1)
                inbound(UBound(inbound)) = theline
                If Left(theline, 1) = "." And foundaddr = False Then
                    foundaddr = True
                    If InStr(1, theline, "MMCI") = 0 And InStr(1, theline, "MLCI") = 0 And InStr(1, theline, "BR") = 0 Then
                        keep = False
                    Else
                        If InStr(1, theline, "CI") > 0 Then
                            inbound(1) = "CAL|" & inbound(1)
                        Else
                            inbound(1) = "EVA|" & inbound(1)
                        End If
                    End If
                End If
            End If
        Wend
        msgin.Close
        Print #1, separator
        Fil.Move ("c:\ARINC\Archives\In\" & currfolder & "\" & Str(seqin) & "_" & Fil.Name)
        '
        ' Now process the inbound msg array
        '
        If keep = True Then
            msg = ""
            For j = 2 To UBound(inbound)
                msg = msg & inbound(j) & "^"
            Next
            tosql = True
            If Left(inbound(1), 3) = "EVA" And InStr(1, msg, "LDM") = 0 Then
                tosql = False
            End If
            If tosql = True Then
                ' put the message into the intranet database
                sql = "INSERT INTO arinc_messages (msgid, msgtxt) VALUES ("
                sql = sql & "'" & inbound(1) & "', "
                sql = sql & "'" & msg & "')"
                dbDataConn.Execute sql
            End If
        End If
    Next
    If found Then
        Print #1, spacer
        Close #1
    End If
    Set fldrIn = Nothing
    Set fldrOut = fso.GetFolder("C:\ARINC\TempPrint")
    found = False
    For Each Fil In fldrOut.Files
        found = True
        seqout = seqout + 1
        headline = "Snd: " & Str(seqout) & " file: " & Fil.Name & " " & FormatDateTime(Now, vbShortDate) & " " & FormatDateTime(Now, vbLongTime) & vbCrLf
'        toline = msgin.ReadLine
'        fromline = msgin.ReadLine
'        If InStr(1, fromline, "QIFNMCA") > 0 Or InStr(1, fromline, "QIFPMCA") > 0 Or InStr(1, fromline, "PEKUDCA") > 0 Then
'            For Each pr In Printers
'                If pr.DeviceName = "CANON" Then Set Printer = pr
'            Next
'        Else
'        End If
'        Printer.Print toline
'        Printer.Print fromline
        Open "\\127.0.0.1\ARINC_OUT" For Output As #2
        Print #2, headline
        Set msgin = Fil.OpenAsTextStream(ForReading)
        While Not msgin.AtEndOfStream
            theline = msgin.ReadLine
            If Len(theline) < 3 Then theline = "  "
            Print #2, theline
        Wend
        msgin.Close
        Print #2, separator
        Fil.Move ("c:\ARINC\Archives\Out\" & currfolder & "\" & Str(seqout) & "_" & Fil.Name)
    Next
    If found Then
        Print #2, spacer
        Close #2
    End If
    Set fldrIn = Nothing
    Set fldrOut = Nothing
    Set fso = Nothing
        
End Sub
