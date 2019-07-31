VERSION 5.00
Begin VB.Form Form9 
   Caption         =   "Give"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form9"
   ScaleHeight     =   12375
   ScaleWidth      =   22800
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   12840
      TabIndex        =   7
      Top             =   2640
      Width           =   2655
   End
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
      Left            =   7680
      TabIndex        =   5
      Top             =   4200
      Width           =   1215
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
      Left            =   9120
      TabIndex        =   4
      Top             =   4200
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   8760
      TabIndex        =   1
      Top             =   2640
      Width           =   2295
   End
   Begin VB.Label Label7 
      Height          =   495
      Left            =   5520
      TabIndex        =   10
      Top             =   5760
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   2520
      TabIndex        =   9
      Top             =   3120
      Width           =   45
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   1080
      TabIndex        =   8
      Top             =   1800
      Width           =   645
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "DATE"
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
      Left            =   13680
      TabIndex        =   6
      Top             =   1920
      Width           =   1050
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "GIVE ATTENDANCE HERE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   6480
      TabIndex        =   3
      Top             =   360
      Width           =   4455
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
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
      Left            =   3960
      TabIndex        =   2
      Top             =   1800
      Width           =   75
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ENTER ROLL NUMBER  WHO IS PRESENT"
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
      Left            =   7440
      TabIndex        =   0
      Top             =   1920
      Width           =   4800
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database
Dim rst As Recordset
Private Sub Command1_Click()
Set db = OpenDatabase("C:\Users\kools\Desktop\college protal\college.mdb")
If Label2.Caption = "SOFTWARE TOOLS" Then
    Set rst = db.OpenRecordset("select * from st")
    Else
    Set rst = db.OpenRecordset("select * from DSA")
End If
    rst.AddNew
    rst.Fields("id") = Text1
    rst.Fields("date") = Text2
    rst.Fields("sem") = Val(Label7.Caption)
    rst.Fields("dept") = Label6.Caption
    rst.Update
    MsgBox " Attendance submitted"
End Sub

Private Sub Command2_Click()
Form8.Show
Unload Form9
End Sub
