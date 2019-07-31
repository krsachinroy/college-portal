VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Dashboard"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   30
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form3"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Height          =   7815
      Left            =   3480
      Picture         =   "Form3.frx":0000
      ScaleHeight     =   7755
      ScaleWidth      =   11475
      TabIndex        =   8
      Top             =   1680
      Width           =   11535
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Exam Cell"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      Left            =   1200
      TabIndex        =   7
      Top             =   6960
      Width           =   1935
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   16920
      TabIndex        =   5
      Top             =   240
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Notice"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      Left            =   15360
      TabIndex        =   4
      Top             =   4680
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Libarary"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1200
      TabIndex        =   2
      Top             =   4320
      Width           =   1995
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Attendance"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   15360
      TabIndex        =   1
      Top             =   1560
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Addmission"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1080
      TabIndex        =   0
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   14760
      TabIndex        =   6
      Top             =   960
      Width           =   75
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "MCKV INSTITUTE OF ENGINEERING"
      Height          =   690
      Left            =   4200
      TabIndex        =   3
      Top             =   480
      Width           =   10290
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Form4.Show
    Form3.Hide
End Sub


Private Sub Command2_Click()
Form7.Show
If Label2.Caption = "" Then
    Form7.Command2.Visible = False
Else
    Form7.Command2.Visible = True
End If
Form3.Hide
End Sub

Private Sub Command3_Click()
If Label2.Caption = "" Then
    MsgBox "First Login your account"
Else
    Form11.Show
    Form3.Hide
End If
End Sub

Private Sub Command4_Click()
Form10.Show
End Sub

Private Sub Command5_Click()
If Command5.Caption = "Login" And Label2.Caption = "" Then
    Form1.Show
    Unload Form3
    Else
     Command5.Caption = "Login"
     Label2.Caption = ""
End If
End Sub


Private Sub Command6_Click()
Form12.Show
Form3.Hide
End Sub
