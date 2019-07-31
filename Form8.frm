VERSION 5.00
Begin VB.Form Form8 
   Caption         =   "Add attendance"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form8"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   "BACK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7800
      TabIndex        =   7
      Top             =   4320
      Width           =   1215
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   9120
      TabIndex        =   6
      Top             =   3000
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SUBMIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      TabIndex        =   4
      Top             =   4320
      Width           =   1215
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   5880
      TabIndex        =   3
      Top             =   3120
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2880
      TabIndex        =   2
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "GIVE ATTENDANCE HERE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   6600
      TabIndex        =   8
      Top             =   480
      Width           =   5625
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Subject"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9360
      TabIndex        =   5
      Top             =   1920
      Width           =   990
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Semester"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6240
      TabIndex        =   1
      Top             =   1920
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Select Department"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   0
      Top             =   1920
      Width           =   2415
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form9.Show
Form9.Label2.Caption = Combo3.Text
Form9.Label5.Caption = "SEMESTER : " + Combo2.Text
Form9.Label6.Caption = Combo1.Text
Form9.Label7.Caption = Combo2.Text
Form9.Label7.Caption = Combo2.Text
Form8.Hide
End Sub

Private Sub Command2_Click()
Form7.Show
Unload Form8
End Sub

Private Sub Form_Load()
Combo1.AddItem "COMPUTER SCIENCE AND ENGINEERING"
Combo1.AddItem "INFORMATION TECHNOLOGY"
Combo2.AddItem "1"
Combo2.AddItem "2"
Combo3.AddItem "DSA"
Combo3.AddItem "SOFTWARE TOOLS"
End Sub
