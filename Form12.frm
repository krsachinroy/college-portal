VERSION 5.00
Begin VB.Form Form12 
   BackColor       =   &H00C0FFC0&
   Caption         =   "exarm "
   ClientHeight    =   8040
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14820
   LinkTopic       =   "Form2"
   ScaleHeight     =   8040
   ScaleWidth      =   14820
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command4 
      Caption         =   "BACK"
      Height          =   495
      Left            =   8160
      TabIndex        =   18
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Add Marks"
      Height          =   495
      Left            =   11760
      TabIndex        =   17
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox Text7 
      Height          =   495
      Left            =   12720
      TabIndex        =   16
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      Height          =   495
      Left            =   10680
      TabIndex        =   15
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   8760
      TabIndex        =   14
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   5040
      TabIndex        =   10
      Top             =   4200
      Width           =   2535
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   5040
      TabIndex        =   8
      Top             =   3480
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "LOG IN"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   6
      Top             =   5160
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   5040
      TabIndex        =   5
      Top             =   2640
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   5040
      TabIndex        =   4
      Top             =   1800
      Width           =   2535
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0FFC0&
      Caption         =   "TEST 3"
      Height          =   495
      Left            =   12720
      TabIndex        =   13
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0FFC0&
      Caption         =   "TEST 2"
      Height          =   495
      Left            =   10680
      TabIndex        =   12
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0FFC0&
      Caption         =   "TEST 1"
      Height          =   495
      Left            =   8760
      TabIndex        =   11
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Subject"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3000
      TabIndex        =   9
      Top             =   4320
      Width           =   810
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "DEPARTMENT"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   7
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "ROLL"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "NAME"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "LOG IN TO SEE MARKS"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   1
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "EXAM CELL"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   615
      Left            =   5040
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database
Dim rst As Recordset
Private Sub Command1_Click()
Dim s1 As Integer, s2 As String, s3 As String, s4 As String
Set db = OpenDatabase("C:\Users\kools\Downloads\EXAM CELL\EXAM CELL\EXAM CELL\STUDENTDATABASE.mdb")
Set rst = db.OpenRecordset("select * from STUDENT")
If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Then
MsgBox "Fill all box correctly"
Else
    Label7.Visible = True
    Label8.Visible = True
    Label9.Visible = True
    Text5.Visible = True
    Text6.Visible = True
    Text7.Visible = True
    While rst.EOF = False
        s2 = rst.Fields("NAME")
        s1 = rst.Fields("ROLL")
        s3 = rst.Fields("DEPARTMENT")
        s4 = rst.Fields("subject")
        If s1 = Text2 Or s2 = Text1 Or s3 = Text3 Or s4 = Text4 Then
            Text5 = rst.Fields("TEST1")
            Text6 = rst.Fields("TEST2")
            Text7 = rst.Fields("TEST3")
        End If
        rst.MoveNext
    Wend
    If Text5 = "" Or Text6 = "" Or Text7 = "" Then
    MsgBox "plz give correct input"
    End If
End If
End Sub

Private Sub Command3_Click()
If Form3.Label2 = "" Then
    MsgBox "plz login first"
    Form1.Show
    Unload Form12
Else
Form13.Show
Unload Form12
End If
End Sub

Private Sub Command4_Click()
Form3.Show
Unload Form12
End Sub

Private Sub Form_Load()
Label7.Visible = False
Label8.Visible = False
Label9.Visible = False
Text5.Visible = False
Text6.Visible = False
Text7.Visible = False
End Sub
