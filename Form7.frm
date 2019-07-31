VERSION 5.00
Begin VB.Form Form7 
   Caption         =   "Attendance"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form7"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   6480
      TabIndex        =   12
      Top             =   4200
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   2880
      TabIndex        =   11
      Top             =   4200
      Width           =   1935
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   11520
      TabIndex        =   8
      Top             =   1800
      Width           =   2175
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   8160
      TabIndex        =   6
      Top             =   1800
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   4800
      TabIndex        =   4
      Top             =   1680
      Width           =   1935
   End
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
      Left            =   16200
      TabIndex        =   3
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ADD ATTENDANCE"
      Height          =   495
      Left            =   15840
      TabIndex        =   2
      Top             =   600
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
      Left            =   16200
      TabIndex        =   1
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "SOFTWARE TOOLS"
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
      Left            =   6120
      TabIndex        =   10
      Top             =   3480
      Width           =   2850
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "DSA"
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
      Left            =   3600
      TabIndex        =   9
      Top             =   3480
      Width           =   630
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "SEMESTER"
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
      Left            =   11760
      TabIndex        =   7
      Top             =   1080
      Width           =   1650
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "DEPARTMENT"
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
      Left            =   8280
      TabIndex        =   5
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "ENTER ROLL NUMBER"
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
      Left            =   4200
      TabIndex        =   0
      Top             =   1080
      Width           =   3165
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database
Dim rst1 As Recordset
Dim rst2 As Recordset
Private Sub Command1_Click()
Label3.Visible = True
Label4.Visible = True
Text2.Visible = True
Text3.Visible = True
Dim s1 As Integer, s2 As Integer, s3 As Integer, s4 As Integer, s5 As String
s1 = 0
s2 = 0
Set db = OpenDatabase("C:\Users\kools\Desktop\college protal\college.mdb")
Set rst1 = db.OpenRecordset("select * from dsa")
Set rst2 = db.OpenRecordset("select * from st")
While rst1.EOF = False
    s5 = rst1.Fields("dept")
    s3 = rst1.Fields("id")
    s4 = rst1.Fields("sem")
    If s5 = Combo1.Text Or s3 = Text1 Or s4 = Val(Combo2.Text) Then
        s1 = s1 + 1
    End If
    rst1.MoveNext
Wend
While rst2.EOF = False
    s5 = rst2.Fields("dept")
    s3 = rst2.Fields("id")
    s4 = rst2.Fields("sem")
    If s5 = Combo1.Text Or s3 = Text1 Or s4 = Val(Combo2.Text) Then
        s2 = s2 + 1
    End If
    rst2.MoveNext
Wend
Text2 = s1
Text3 = s2
End Sub

Private Sub Command2_Click()
Form8.Show
End Sub

Private Sub Command3_Click()
Form3.Show
Unload Form7
End Sub

Private Sub Form_Load()
Label3.Visible = False
Label4.Visible = False
Text2.Visible = False
Text3.Visible = False
Combo1.AddItem "COMPUTER SCIENCE AND ENGINEERING"
Combo1.AddItem "INFORMATION TECHNOLOGY"
Combo2.AddItem "1"
Combo2.AddItem "2"
End Sub
