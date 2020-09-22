VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Send E-Mail, directly through VB."
   ClientHeight    =   4470
   ClientLeft      =   150
   ClientTop       =   675
   ClientWidth     =   4875
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   4875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock SMTP 
      Left            =   4440
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   4575
      Begin VB.TextBox txtSub 
         Height          =   285
         Left            =   840
         TabIndex        =   8
         Text            =   "Subject Here"
         Top             =   360
         Width           =   3615
      End
      Begin VB.TextBox txtTo 
         Height          =   285
         Left            =   840
         TabIndex        =   7
         Text            =   "Whoever@whatever.com"
         Top             =   720
         Width           =   3615
      End
      Begin VB.TextBox txtFrom 
         Height          =   285
         Left            =   840
         TabIndex        =   6
         Text            =   "Your E-Mail, name, it's up to you."
         Top             =   1080
         Width           =   3615
      End
      Begin VB.TextBox txtServer 
         Height          =   285
         Left            =   840
         TabIndex        =   5
         Text            =   "something like... mail.hotmail.com usually works."
         Top             =   1440
         Width           =   3615
      End
      Begin VB.Label Label1 
         Caption         =   "Subject:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "To:"
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   11
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "From:"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   10
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Server:   "
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1440
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      Top             =   3960
      Width           =   735
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Send"
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   3960
      Width           =   735
   End
   Begin VB.TextBox txtBody 
      Height          =   1575
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "Form1.frx":0442
      Top             =   2280
      Width           =   4575
   End
   Begin VB.Label lblStats 
      Caption         =   "Ready....."
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   4080
      Width           =   2655
   End
   Begin VB.Label Label3 
      Caption         =   "Message Body:"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&Information"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


' contain the info neede to send E_Mail
Dim sTo      As String
Dim sFrom    As String
Dim sSubject As String
Dim sServer  As String
Dim sBody    As String



Private Sub cmdCancel_Click()
    Unload Me
    End
End Sub


' once all needed inf is gathered send the mail in one line of code.
Private Sub cmdSend_Click()
    Mail sSubject, sTo, sFrom, sBody, sServer
End Sub

Private Sub mnuAbout_Click()
   MsgBox " I wasn't sure if you wanted to send mail through a visible interface, or through code." & _
          " SO I made this project accomplish both tasks with minimal work. After you change the form and winsock control to a more suiting name, you should be able to send e_mail with one line of code..... I hope. Please let me know of any problems. " & vbCrLf & vbCrLf & "    gh0ul"
          
End Sub


'
' when
Private Sub SMTP_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next

    Dim datad As String
    SMTP.GetData datad, vbString
    LastSMTP = datad
End Sub


'
' notice I did no Error checking, that means if you don't enter
' the info right, it may not work.
'

Private Sub txtBody_Change()
    sBody = txtBody
End Sub

Private Sub txtFrom_Change()
   sFrom = txtFrom
End Sub

Private Sub txtServer_Change()
    sServer = txtServer
End Sub

Private Sub txtSub_Change()
    sSubject = txtSub
End Sub

Private Sub txtTo_Change()
    sTo = txtTo
End Sub
