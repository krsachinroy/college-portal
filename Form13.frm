VERSION 5.00
Begin VB.Form Form13 
   BackColor       =   &H00C0FFC0&
   Caption         =   "add marks"
   ClientHeight    =   7935
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14025
   LinkTopic       =   "Form1"
   ScaleHeight     =   7935
   ScaleWidth      =   14025
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text7 
      Height          =   495
      Left            =   2040
      TabIndex        =   19
      Top             =   4440
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
      Caption         =   "BACK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      TabIndex        =   17
      Top             =   6480
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "CLEAN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8880
      TabIndex        =   16
      Top             =   5520
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "ADD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   15
      Top             =   5640
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      DataField       =   "TEST 3"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   8760
      TabIndex        =   14
      Top             =   3720
      Width           =   1335
   End
   Begin VB.TextBox Text5 
      DataField       =   "TEST 2"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   8760
      TabIndex        =   13
      Top             =   2880
      Width           =   1335
   End
   Begin VB.TextBox Text4 
      DataField       =   "TEST 1"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   8760
      TabIndex        =   12
      Top             =   2040
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      DataField       =   "DEPARTMENT"
      DataSource      =   "Data1"
      Height          =   405
      Left            =   2040
      TabIndex        =   7
      Top             =   3720
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      DataField       =   "ROLL"
      DataSource      =   "Data1"
      Height          =   405
      Left            =   2040
      TabIndex        =   6
      Top             =   2880
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      DataField       =   "NAME"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C0FFC0&
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
      ForeColor       =   &H80000001&
      Height          =   375
      Left            =   360
      TabIndex        =   18
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "TEST 3"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7080
      TabIndex        =   11
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "TEST 2"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7080
      TabIndex        =   10
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "TEST 1"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7080
      TabIndex        =   9
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "MARKS"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   7920
      TabIndex        =   8
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "STUDENT INFORMATION"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "DEPARTMENT"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "ROLL"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "NAME"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "EXAM CELL"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   4920
      TabIndex        =   0
      Top             =   480
      Width           =   2295
   End
End
Attribute VB_Name = "Form13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim db As Database
Dim rst As Recordset
Private Sub Command1_Click()
Set db = OpenDatabase("C:\Users\kools\Downloads\EXAM CELL\EXAM CELL\EXAM CELL\STUDENTDATABASE.mdb")
Set rst = db.OpenRecordset("select * from STUDENT")
rst.AddNew
rst.Fields("NAME") = Text1
rst.Fields("ROLL") = Text2
rst.Fields("DEPARTMENT") = Text3
rst.Fields("TEST1") = Text4
rst.Fields("TEST2") = Text5
rst.Fields("TEST3") = Text6
rst.Fields("subject") = Text7
rst.Update
End Sub

Private Sub Command4_Click()
Text1 = ""
Text2 = ""
Text3 = ""
Text4 = ""
Text5 = ""
Text6 = ""
Text7 = ""
End Sub

Private Sub Command5_Click()
Form12.Show
Unload Form13
End Sub

Private Sub Form_Load()

End Sub
