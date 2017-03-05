Attribute VB_Name = "modFTIBasics"

' ***************************************************************************
'**                             FTI Basics Module                           **
'**                                 Version 6                               **
'**              (C) Copyright 1996, 1999, 2002, 2006, 2009, 2011           **
'**                          Fullman Technologies, Inc.                     **
'**                             All rights reserved.                        **
' ***************************************************************************
Sub CheckForHoliday()

    '--------------------------------------------------------------------------------------------------------
    ' This routine checks the Holiday Database Table to determine if the market is opened or closed for a
    ' holiday. If the market is closed the program will terminate and not perform any analysis or transmit
    ' any reports.
    '
    ' Constructed by Scott H. Fullman, Fullman Technologies, Inc., on Jan. 22, 2006.
    '--------------------------------------------------------------------------------------------------------
    
    Dim sTodayDate      As String * 10
    
    Dim cnnl As New ADODB.Connection
    Dim rstOptions As ADODB.Recordset
    Set cnnl = New ADODB.Connection
    Dim cmdChange       As ADODB.Command
    cnnl.Open "DSN=FTIBasics"
    cnnl.CommandTimeout = 100
    Set rstOptions = New ADODB.Recordset
    Set cmdChange = New ADODB.Command
    Set cmdChange.ActiveConnection = cnnl
    rstOptions.CursorType = adOpenDynamic
    
    rstOptions.LockType = adLockOptimistic
    
'    rstOptions.Open "Holiday", cnnl, , , adCmdTable
    
    sTodayDate = Format(Date, "mm/dd/yyyy")
    
    cmdChange.CommandText = "Select * FROM Holidays Where Holiday=#" & sTodayDate & "#;"


    Set rstOptions = cmdChange.Execute
    
    If rstOptions.BOF Or rstOptions.EOF Then
        Exit Sub
    Else
        End
    End If
    
End Sub
Sub GetDisclaimer(sDisclaimerName As String, sDisclaimer As String)

    '-----------------------------------------------------------------------------------------------------------
    ' This routine checks the Disclaimer table of the FTIBasics database for the requested disclaimer.
    '
    ' Modified by Scott H Fullman, Fullman Technologies, Inc. on 12/31/2013.
    '-----------------------------------------------------------------------------------------------------------
    
    Dim cnnl As New ADODB.Connection
    Dim rstOptions As ADODB.Recordset
    Set cnnl = New ADODB.Connection
    Dim cmdChange       As ADODB.Command
    cnnl.Open "DSN=FTIBasics"
    cnnl.CommandTimeout = 100
    Set rstOptions = New ADODB.Recordset
    Set cmdChange = New ADODB.Command
    Set cmdChange.ActiveConnection = cnnl
    rstOptions.CursorType = adOpenDynamic
    
    rstOptions.LockType = adLockOptimistic
    
    cmdChange.CommandText = "Select * From Disclaimers Where DisclaimerName='" & sDisclaimerName & "';"
    Set rstOptions = cmdChange.Execute
    
    sDisclaimer = rstOptions!Disclaimer
    
End Sub
Function NameCheck(sOldName As String) As String

    '------------------------------------------------------------------------------------
    ' This function checks the company name (sOldName) for odd characters and converts them
    ' into Windows acceptable symbols for the filenames.
    '
    ' Added 6/7/2011
    '------------------------------------------------------------------------------------
    
    Dim iLocation   As Integer
    Dim iLength     As Integer
    Dim iCount      As Integer
    
    Dim sTempLeft   As String
    Dim sTempRight  As String
    
    sOldName = Trim(sOldName)
    
    iLength = Len(sOldName)
  '  MsgBox ">>" & sOldName
    
    For iCount = 1 To iLength
    
        If Mid(sOldName, iCount, 1) = "'" Then
            sTempLeft = Left(sOldName, iCount - 1)
            sTempRight = Right(sOldName, iLength - (iCount))
            sOldName = sTempLeft & "`" & sTempRight
            Exit For
        Else
            ' Nothing else
        End If
        
        
        
    Next
    
    NameCheck = sOldName

End Function
Function SymbolFix(sOldSymbol As String) As String

    '------------------------------------------------------------------------------------
    ' This function checks the symbol (sOldSymbol) for odd characters and converts them
    ' into Windows acceptable symbols for the filenames.
    '
    ' Added 6/7/2011
    '------------------------------------------------------------------------------------
    
    Dim iLocation   As Integer
    Dim iLength     As Integer
    Dim iCount      As Integer
    
    Dim sTempLeft   As String
    Dim sTempRight  As String
    
    sOldSymbol = Trim(sOldSymbol)
    
    iLength = Len(sOldSymbol)
    
    For iCount = 1 To iLength
    
        If Mid(sOldSymbol, iCount, 1) = "/" Then
            sTempLeft = Left(sOldSymbol, iCount - 1)
            sTempRight = Right(sOldSymbol, iLength - (iCount - 1))
            sOldSymbol = sTempLeft & "_" & sTempRight
        End If
        
        If Mid(sOldSymbol, iCount, 1) = "'" Then
            sTempLeft = Left(sOldSymbol, iCount - 1)
            sTempRight = Right(sOldSymbol, iLength - (iCount - 1))
            sOldSymbol = sTempLeft & "`" & sTempRight
        End If
        
    Next
    
    SymbolFix = sOldSymbol
    
End Function

Sub SendReport(sHTML As String, sSubject As String, sReportName As String, sAttach As String)

    '----------------------------------------------------------------------------------
    ' This routine takes the report generated and sends it to the distribution list
    ' on the FTI MailGen database.
    '----------------------------------------------------------------------------------
    
    Dim sMailTo             As String
    
    
    Dim cnnl As New ADODB.Connection
    Dim rstOptions As ADODB.Recordset
    Dim rstSend    As ADODB.Recordset
    Set cnnl = New ADODB.Connection
    Dim cmdChange       As ADODB.Command
    Dim cmdChangeSend   As ADODB.Command
    cnnl.Open "DSN=MailGen"
    cnnl.CommandTimeout = 100
    Set rstOptions = New ADODB.Recordset
    Set rstSend = New ADODB.Recordset
    Set cmdChange = New ADODB.Command
    Set cmdChange.ActiveConnection = cnnl
    Set cmdChangeSend = New ADODB.Command
    Set cmdChangeSend.ActiveConnection = cnnl
    rstOptions.CursorType = adOpenDynamic
    rstOptions.LockType = adLockOptimistic
    rstSend.CursorType = adOpenDynamic
    rstSend.LockType = adLockOptimistic
    
    cmdChange.CommandText = "Select * From MailDistList Where  " & sReportName & "=1;"
    'MsgBox cmdChange.CommandText
    Set rstOptions = cmdChange.Execute
    
    rstOptions.MoveFirst
    
        Dim cdoMsg As New CDO.Message
    Dim cdoConf As New CDO.Configuration
    Dim Flds
    
    Const cdoSendUsingPort = 2
    
    Set cdoMsg = CreateObject("CDO.Message")
    Set cdoConf = CreateObject("CDO.Configuration")
    
    Set Flds = cdoConf.Fields
        
    With Flds
        .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 '
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp-server.si.rr.com"
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = cdoBasic
        .Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "sfullman"
        .Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "j1ngles1"
    .Update
    End With
 ' MsgBox sAttach
    If Trim(sAttach) <> "" Then
        cdoMsg.AddAttachment sAttach
    End If
    
    Do While Not rstOptions.EOF
        DoEvents
    
    ' Apply the settings to the message.
        With cdoMsg
            Set .Configuration = cdoConf
            .To = rstOptions!UserID
            '.CC = "sfullman@si.rr.com"
            .From = "Increasing Alpha<sfullman@si.rr.com>"
            .Subject = sSubject
            .HTMLBody = sHTML

    '        If strCc <> "" Then .CC = strCc
    '        If strBcc <> "" Then .BCC = strBcc
            .Send
    
            
        End With
    
        cmdChangeSend.CommandText = "Insert Into ReportsSent (UserID,  LastName, FirstName,Company,ReportName, ReportSentDate, ReportSentTime) Values ('"
        cmdChangeSend.CommandText = cmdChangeSend.CommandText & rstOptions!UserID & "','"
        cmdChangeSend.CommandText = cmdChangeSend.CommandText & rstOptions!LastName & "','"
        cmdChangeSend.CommandText = cmdChangeSend.CommandText & rstOptions!FirstName & "','"
        cmdChangeSend.CommandText = cmdChangeSend.CommandText & rstOptions!Company & "','"
        cmdChangeSend.CommandText = cmdChangeSend.CommandText & sReportName & "',#"
        cmdChangeSend.CommandText = cmdChangeSend.CommandText & Format(Date, "mm/dd/yyyy") & "#,#"
        cmdChangeSend.CommandText = cmdChangeSend.CommandText & Format(Time, "hh:mm:ss") & "#);"
        Set rstSend = cmdChangeSend.Execute
        rstOptions.MoveNext
    
    Loop
        
End Sub

Sub SendCustomReport(sHTML As String, sSubject As String, sReportName As String, sAttach As String, sEMailAddress As String, sCCAddress As String)

    '----------------------------------------------------------------------------------
    ' This routine takes the report generated and sends it to the distribution list
    ' on the FTI MailGen database.
    '----------------------------------------------------------------------------------
    
    Dim sMailTo             As String
    
    
    Dim cnnl As New ADODB.Connection
    Dim rstOptions As ADODB.Recordset
    Dim rstSend    As ADODB.Recordset
    Set cnnl = New ADODB.Connection
    Dim cmdChange       As ADODB.Command
    Dim cmdChangeSend   As ADODB.Command
    cnnl.Open "DSN=MailGen"
    cnnl.CommandTimeout = 100
    Set rstOptions = New ADODB.Recordset
    Set rstSend = New ADODB.Recordset
    Set cmdChange = New ADODB.Command
    Set cmdChange.ActiveConnection = cnnl
    Set cmdChangeSend = New ADODB.Command
    Set cmdChangeSend.ActiveConnection = cnnl
    rstOptions.CursorType = adOpenDynamic
    rstOptions.LockType = adLockOptimistic
    rstSend.CursorType = adOpenDynamic
    rstSend.LockType = adLockOptimistic
    
    
    Dim cdoMsg As New CDO.Message
    Dim cdoConf As New CDO.Configuration
    Dim Flds
    
    Const cdoSendUsingPort = 2
    
    'Set cdoMsg =  CreateObject("CDO.Message")
    'Set cdoConf = CreateObject("CDO.Configuration")
    
    Set Flds = cdoConf.Fields
        
    With Flds
        .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 '
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp-server.si.rr.com"
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = cdoBasic
        .Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "sfullman"
        .Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "j1ngles1"
    .Update
    End With
    
    ' Apply the settings to the message.
        With cdoMsg
            Set .Configuration = cdoConf
            .To = sEMailAddress
            If Trim(sCCAddress) <> "" Then
                .CC = sCCAddress
            End If
            .From = "Increasing Alpha<sfullman@si.rr.com>"
            .Subject = sSubject
            .HTMLBody = sHTML
            If Trim(sAttach) <> "" Then
                .AddAttachment sAttach
            End If
    '        If strCc <> "" Then .CC = strCc
    '        If strBcc <> "" Then .BCC = strBcc
            .Send
    
            
        End With
    
'        cmdChangeSend.CommandText = "Insert Into ReportsSent (UserID,  LastName, FirstName,Company,ReportName, ReportSentDate, ReportSentTime) Values ('"
'        cmdChangeSend.CommandText = cmdChangeSend.CommandText & rstOptions!UserID & "','"
'        cmdChangeSend.CommandText = cmdChangeSend.CommandText & rstOptions!LastName & "','"
'        cmdChangeSend.CommandText = cmdChangeSend.CommandText & rstOptions!FirstName & "','"
'        cmdChangeSend.CommandText = cmdChangeSend.CommandText & rstOptions!Company & "','"
'        cmdChangeSend.CommandText = cmdChangeSend.CommandText & sReportName & "',#"
'        cmdChangeSend.CommandText = cmdChangeSend.CommandText & Format(Date, "mm/dd/yyyy") & "#,#"
'        cmdChangeSend.CommandText = cmdChangeSend.CommandText & Format(Time, "hh:mm:ss") & "#);"
'        Set rstSend = cmdChangeSend.Execute
    
End Sub

Sub SendErrorMessage(sHTML As String)


    '----------------------------------------------------------------------------------
    ' This routine takes the report generated and sends it to the distribution list
    ' on the FTI MailGen database.
    '----------------------------------------------------------------------------------
    
    Dim sMailTo             As String
    
    
    Dim cnnl As New ADODB.Connection
    Dim rstOptions As ADODB.Recordset
    Dim rstSend    As ADODB.Recordset
    Set cnnl = New ADODB.Connection
    Dim cmdChange       As ADODB.Command
    Dim cmdChangeSend   As ADODB.Command
    cnnl.Open "DSN=MailGen"
    cnnl.CommandTimeout = 100
    Set rstOptions = New ADODB.Recordset
    Set rstSend = New ADODB.Recordset
    Set cmdChange = New ADODB.Command
    Set cmdChange.ActiveConnection = cnnl
    Set cmdChangeSend = New ADODB.Command
    Set cmdChangeSend.ActiveConnection = cnnl
    rstOptions.CursorType = adOpenDynamic
    rstOptions.LockType = adLockOptimistic
    rstSend.CursorType = adOpenDynamic
    rstSend.LockType = adLockOptimistic
    
    
    Dim cdoMsg As New CDO.Message
    Dim cdoConf As New CDO.Configuration
    Dim Flds
    
    Const cdoSendUsingPort = 2
    
    'Set cdoMsg =  CreateObject("CDO.Message")
    'Set cdoConf = CreateObject("CDO.Configuration")
    
    Set Flds = cdoConf.Fields
        
    With Flds
        .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 '
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp-server.si.rr.com"
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = cdoBasic
        .Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "sfullman"
        .Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "j1ngles1"
    .Update
    End With
    
    ' Apply the settings to the message.
        With cdoMsg
            Set .Configuration = cdoConf
            .To = "sfullman@fullmantech.com"
            .CC = "9176277409@vztext.com"
            .From = "Increasing Alpha<sfullman@si.rr.com>"
            .Subject = sSubject
            .HTMLBody = sHTML

            .Send
    
            
        End With
End Sub
