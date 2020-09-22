Attribute VB_Name = "modMail"
Option Explicit

Public LastSMTP As String

Public Declare Function GetTickCount Lib "kernel32" () As Long

'---------------------------------------------------------------------------
' AUTHOR: gh0ul
'
' PROCEDURE NAME: Mail
' PURPOSE:        calls a function to initialze mail
' PARAMETERS:     SUbject, To, From, Host, Body
'
' RETURNS:        nothing
'
'---------------------------------------------------------------------------
' DATE:  September,30 99
' TIME:  01:01
'---------------------------------------------------------------------------

Sub Mail(strSubject As String, strTo As String, strFrom As String, _
             strBody As String, strHost As String)
             
    Call InitMail(strSubject, strTo, strFrom, strBody, strHost)
    
End Sub

'---------------------------------------------------------------------------
' AUTHOR: gh0ul
'
' PROCEDURE NAME: InitMail()
' PURPOSE:        this sub attempts to send the message.
' PARAMETERS:     same as above
'
' RETURNS:        nada
'
'---------------------------------------------------------------------------
' DATE:  September,30 99
' TIME:  01:03
'---------------------------------------------------------------------------

Sub InitMail(strSubject As String, strTo As String, strFrom As String, _
             strBody As String, strHost As String)


Test:
    DoEvents
    Dim Res
    Debug.Print "Trying to send mail " & Timer
    Res = doMail(strSubject, strTo, strFrom, strBody, strHost)
    If Res = True Then GoTo Wait
    If Res = False Then GoTo Test

Wait:
Form1.lblStats = "Mail sent " & Timer


End Sub

'---------------------------------------------------------------------------
' AUTHOR: gh0ul
'
' PROCEDURE NAME: doMail()
' PURPOSE:        Sends the mail via winsock... connects, sends, disconnects.
' PARAMETERS:     same as above
'
' RETURNS:        Boolean.... this function is called until true, else mail
'                 send will fail
'
'---------------------------------------------------------------------------
' DATE:  September,30 99
' TIME:  01:04
'---------------------------------------------------------------------------

Public Function doMail(strSubject As String, strTo As String, strFrom As String, _
             strBody As String, strHost As String) As Boolean

    On Error Resume Next
    Dim CTimer As Long
    Dim Server As String
    Dim UserName As String

    LastSMTP = ""

    Randomize Timer

    Server = strHost

    Form1.lblStats = "Trying to connect to " & Server

    With Form1.SMTP
       .Close
       .LocalPort = 0
       .RemoteHost = Server
       .RemotePort = 25  ' this usually works for e_Mail
       .Connect
    End With

    CTimer = Timer
    Dim dbgState As Integer
    dbgState = 10
    Do
        If Len(LastSMTP) > 1 Then GoTo SendMail
        If Form1.SMTP.State <> dbgState Then
          Form1.lblStats = Form1.SMTP.State
            dbgState = Form1.SMTP.State
        If Form1.SMTP.State = 9 Then Exit Do
        End If
        DoEvents

    Loop Until CTimer + 30 < Timer

    doMail = False
    Form1.lblStats = "Timed Out..."
    Form1.lblStats = "Last SMTP: " & LastSMTP

    Exit Function

SendMail:

    Pause 0.5


    With Form1
        .SMTP.SendData "HELO " & String(256, "A") & vbCrLf 'hide ip from old sendmail
        .SMTP.SendData "MAIL FROM:" & strFrom & "@" & Form1.SMTP.LocalIP & vbCrLf
        .SMTP.SendData "RCPT TO:" & strTo & vbCrLf
        .SMTP.SendData "RCPT TO" & strTo & vbCrLf
        .SMTP.SendData "DATA" & vbCrLf
        
        Pause 0.5

        .SMTP.SendData "TO: " & strTo & vbCrLf
        .SMTP.SendData "FROM: " & LCase(strFrom) & "@" & Form1.SMTP.LocalIP & vbCrLf
        .SMTP.SendData "Subject: " & strSubject & vbCrLf
        .SMTP.SendData vbCrLf
        .SMTP.SendData String(5, Chr(13)) & vbCrLf

        Pause 0.5

        .SMTP.SendData "Time Sent:    " & Time & vbCrLf & "IP Address:    " & Form1.SMTP.LocalIP & vbCrLf
        .SMTP.SendData vbCrLf & strBody & vbCrLf
        .SMTP.SendData vbCrLf
        .SMTP.SendData "." & vbCrLf
    End With
    
        CTimer = Timer
        Form1.lblStats = "Email Sent to " & strTo
        
   
   
    Do
       DoEvents
    Loop Until CTimer + 20 < Timer

    With Form1.SMTP
      .Close
      .LocalPort = 0
    End With
    
    doMail = True

    Form1.lblStats = "Closing Connection..."
End Function



' you need slight pauses when sending multiple strings of data with
' winsock. this function does that in 1000th of seconds  1000 = 1 sec
Sub Pause(HowLong As Long)
    '
    Dim u%, Tick As Long
    
    Tick = GetTickCount
    
    Do
      u% = DoEvents
    Loop Until Tick + HowLong < GetTickCount
End Sub
