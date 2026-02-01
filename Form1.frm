VERSION 5.00
Object = "{2253C784-C925-4714-A228-DB76DB6E7C4A}#1.0#0"; "OscSeven.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Oscsev OCX Example "
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6000
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   6000
   StartUpPosition =   3  'Windows Default
   Begin OscSeven.OscSev OscSev1 
      Height          =   495
      Left            =   240
      TabIndex        =   26
      Top             =   3480
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   873
   End
   Begin RichTextLib.RichTextBox Message 
      Height          =   2415
      Left            =   120
      TabIndex        =   25
      Top             =   4080
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   4260
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"Form1.frx":0000
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   5280
      TabIndex        =   20
      Text            =   "0"
      Top             =   3480
      Width           =   495
   End
   Begin VB.Frame Frame1 
      Caption         =   "Main"
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5775
      Begin VB.Frame Frame2 
         Caption         =   "Functions"
         Height          =   1935
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   5535
         Begin VB.CommandButton Command14 
            Caption         =   "Idle"
            Height          =   375
            Left            =   4440
            TabIndex        =   19
            Top             =   1320
            Width           =   615
         End
         Begin VB.CommandButton Command13 
            Caption         =   "Go Invisible"
            Height          =   375
            Left            =   3120
            TabIndex        =   18
            Top             =   1320
            Width           =   1335
         End
         Begin VB.CommandButton Command12 
            Caption         =   "Away"
            Height          =   375
            Left            =   2400
            TabIndex        =   17
            Top             =   1320
            Width           =   735
         End
         Begin VB.CommandButton Command11 
            Caption         =   "Leave"
            Height          =   375
            Left            =   4440
            TabIndex        =   16
            Top             =   960
            Width           =   615
         End
         Begin VB.CommandButton Command10 
            Caption         =   "Send Chat Msg"
            Height          =   375
            Left            =   3120
            TabIndex        =   15
            Top             =   960
            Width           =   1335
         End
         Begin VB.CommandButton Command9 
            Caption         =   "Join ch"
            Height          =   375
            Left            =   2400
            TabIndex        =   14
            Top             =   960
            Width           =   735
         End
         Begin VB.CommandButton Command8 
            Caption         =   "Game"
            Height          =   375
            Left            =   4440
            TabIndex        =   13
            Top             =   600
            Width           =   615
         End
         Begin VB.CommandButton Command7 
            Caption         =   "Buddylist Send"
            Height          =   375
            Left            =   3120
            TabIndex        =   12
            Top             =   600
            Width           =   1335
         End
         Begin VB.CommandButton Command6 
            Caption         =   "Send DC"
            Height          =   375
            Left            =   2400
            TabIndex        =   11
            Top             =   600
            Width           =   735
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Talk"
            Height          =   375
            Left            =   4440
            TabIndex        =   10
            Top             =   240
            Width           =   615
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Send Expression"
            Height          =   375
            Left            =   3120
            TabIndex        =   9
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Send IM"
            Height          =   375
            Left            =   2400
            TabIndex        =   8
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox Text4 
            Height          =   285
            Left            =   120
            TabIndex        =   7
            Text            =   "This is a Test!!"
            Top             =   600
            Width           =   2175
         End
         Begin VB.TextBox Text3 
            Height          =   285
            Left            =   120
            TabIndex        =   6
            Text            =   "Jackass"
            Top             =   240
            Width           =   2175
         End
         Begin VB.Label Label1 
            Caption         =   "This is a small example. You should be able to figure out the rest!!!!"
            ForeColor       =   &H000000C0&
            Height          =   735
            Left            =   120
            TabIndex        =   21
            Top             =   1080
            Width           =   2055
         End
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Log out"
         Height          =   375
         Left            =   2760
         TabIndex        =   4
         Top             =   720
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Log In"
         Height          =   375
         Left            =   2760
         TabIndex        =   3
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   240
         PasswordChar    =   "+"
         TabIndex        =   2
         Text            =   "1234"
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   240
         TabIndex        =   1
         Text            =   "Slapbot"
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label4 
         Caption         =   "------------"
         Height          =   255
         Left            =   3720
         TabIndex        =   24
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "0"
         Height          =   255
         Left            =   3720
         TabIndex        =   23
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Status:......"
         Height          =   255
         Left            =   3720
         TabIndex        =   22
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "----- IN Coming IM ------"
      Height          =   255
      Left            =   2160
      TabIndex        =   27
      Top             =   3600
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
OscSev1.Login Text1, Text2, "login.oscar.aol.com", 1  '<---- 1 for mobile 2 for normal
End Sub

Private Sub Command10_Click()
OscSev1.sendChatMessage Text5, Text4
End Sub

Private Sub Command11_Click()
OscSev1.endSession Text5
End Sub

Private Sub Command12_Click()
OscSev1.SetAway Text4  '<--- to come back use just enter ("")
End Sub

Private Sub Command13_Click()
OscSev1.SetInvisable False
End Sub

Private Sub Command14_Click()
OscSev1.unIdle
End Sub

Private Sub Command2_Click()
OscSev1.Loggoff
End Sub

Private Sub Command3_Click()
OscSev1.SendMessage Text3, Text4, False
End Sub

Private Sub Command4_Click()
OscSev1.ThemeSend Text3, Text4, "evil"
End Sub

Private Sub Command5_Click()
OscSev1.Sendtalk Text3
End Sub

Private Sub Command6_Click()
OscSev1.SendConnect Text3
End Sub

Private Sub Command7_Click()
OscSev1.SendBuddylist Text3, Text4, 50
End Sub

Private Sub Command8_Click()
OscSev1.Sendgame Text3, "www.sevenz.net", Text4, Text4
End Sub

Private Sub Command9_Click()
OscSev1.JoinChat "slap", 4
End Sub


Private Sub Message_Change()
On Error Resume Next
    Message.SelLength = 0


    If Len(Message.Text) > 0 Then


        If Right$(Message.Text, 1) = vbCrLf Then
         Message.SelStart = Len(Message.Text) - 1
            Exit Sub
        End If
       Message.SelStart = Len(Message.Text)
    End If
End Sub

Private Sub OscSev1_chatReady(strChannel As String, intIndex As Integer)
Text5 = intIndex
End Sub

Private Sub OscSev1_FunctionSent()
Label3 = Label3 + 1
End Sub

Private Sub OscSev1_incomingIM(strName As String, strMessage As String)
'Text6.Text = strName & ": " & strMessage

Call RTFUpdate(Message, "\par\plain\fs16\cf1\b " & CStr(strName) & ": \plain\fs16\cf0 " & FixRTF(KillHTML(CStr(strMessage))))

End Sub

Private Sub OscSev1_loggedIn(strFormattedName As String, strEmailAddy As String)
Label2.Caption = "Status: " & strFormattedName
End Sub

Private Sub OscSev1_loggedOut(strFormattedName As String)
Label1.Caption = "Status:......"
End Sub

Private Sub OscSev1_NotOnline()
Label4.Caption = "User Not Online"
End Sub
