VERSION 5.00
Begin VB.Form Form6 
   AutoRedraw      =   -1  'True
   Caption         =   "student details"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form6"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "STUDENT DETAILS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8655
      Left            =   1560
      TabIndex        =   0
      Top             =   720
      Width           =   15495
      Begin VB.ListBox List7 
         Height          =   4350
         Left            =   240
         TabIndex        =   15
         Top             =   2160
         Width           =   1455
      End
      Begin VB.ListBox List6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4140
         Left            =   4200
         TabIndex        =   13
         Top             =   2160
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Back"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   6240
         TabIndex        =   11
         Top             =   7080
         Width           =   1095
      End
      Begin VB.ListBox List5 
         Height          =   3960
         Left            =   12120
         TabIndex        =   10
         Top             =   2160
         Width           =   3015
      End
      Begin VB.ListBox List4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4140
         Left            =   10320
         TabIndex        =   8
         Top             =   2160
         Width           =   1455
      End
      Begin VB.ListBox List3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4140
         Left            =   8280
         TabIndex        =   6
         Top             =   2160
         Width           =   1575
      End
      Begin VB.ListBox List2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4140
         Left            =   6240
         TabIndex        =   4
         Top             =   2160
         Width           =   1575
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4380
         Left            =   1920
         TabIndex        =   2
         Top             =   2160
         Width           =   1695
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "ROLL NUMBER"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   14
         Top             =   1440
         Width           =   1440
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Course"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4440
         TabIndex        =   12
         Top             =   1320
         Width           =   765
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   13200
         TabIndex        =   9
         Top             =   1320
         Width           =   885
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "E-mail"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   10680
         TabIndex        =   7
         Top             =   1320
         Width           =   660
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Phone number"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   8160
         TabIndex        =   5
         Top             =   1320
         Width           =   1560
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Department"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6360
         TabIndex        =   3
         Top             =   1320
         Width           =   1275
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Students Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1920
         TabIndex        =   1
         Top             =   1320
         Width           =   1665
      End
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database
Dim rst As Recordset
Dim i As Integer
Private Sub Command1_Click()
Form4.Show
Unload Form6
End Sub

Private Sub Form_Load()
Set db = OpenDatabase("C:\Users\kools\Desktop\college protal\college.mdb")
Set rst = db.OpenRecordset("select * from admission")
While rst.EOF = False
    List7.AddItem rst.Fields("roll")
    List7.AddItem "-----------------"
    List1.AddItem rst.Fields("name")
    List1.AddItem "-----------------"
    List2.AddItem rst.Fields("department")
    List2.AddItem "-----------------"
    List6.AddItem rst.Fields("course")
    List6.AddItem "-----------------"
    List3.AddItem rst.Fields("phone")
    List3.AddItem "-----------------"
    List4.AddItem rst.Fields("email")
    List4.AddItem "-----------------"
    List5.AddItem rst.Fields("address")
    List5.AddItem "-----------------"
    rst.MoveNext
Wend
End Sub
