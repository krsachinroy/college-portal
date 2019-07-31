VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Admission information"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form4"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
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
      Left            =   10680
      TabIndex        =   7
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "STUDENT LIST"
      Height          =   495
      Left            =   15600
      TabIndex        =   6
      Top             =   600
      Width           =   1455
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   930
      Left            =   12000
      TabIndex        =   3
      Top             =   1800
      Width           =   4455
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   930
      Left            =   3600
      TabIndex        =   1
      Top             =   1800
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "TAKE ADDMISSION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7800
      TabIndex        =   4
      Top             =   3720
      Width           =   1815
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "ADMISSION INFORMATION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   5280
      TabIndex        =   5
      Top             =   360
      Width           =   7665
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "FEES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   12120
      TabIndex        =   2
      Top             =   1200
      Width           =   945
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "COURSES"
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
      Left            =   3720
      TabIndex        =   0
      Top             =   1200
      Width           =   1485
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form5.Show
Unload Form4
End Sub

Private Sub Command2_Click()
If Form3.Label2.Caption = "" Then
    MsgBox "First login your authorized account"
    Form1.Show
    Form4.Hide
Else
    Form6.Show
    Unload Form4
End If
End Sub

Private Sub Command3_Click()
Form3.Show
Unload Form4
End Sub

Private Sub Form_Load()
List1.AddItem "B.TECH"
List1.AddItem "M.TECH"
List2.AddItem "Rs.40000 per semester"
List2.AddItem "Rs.80000 per semester"
End Sub
