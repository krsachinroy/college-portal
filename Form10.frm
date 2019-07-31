VERSION 5.00
Begin VB.Form Form10 
   Caption         =   "NOTICE"
   ClientHeight    =   7335
   ClientLeft      =   1515
   ClientTop       =   2055
   ClientWidth     =   11880
   LinkTopic       =   "Form2"
   ScaleHeight     =   7335
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command8 
      Caption         =   "Add New Notice"
      Height          =   495
      Left            =   10080
      TabIndex        =   10
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Delete"
      Height          =   735
      Left            =   2520
      TabIndex        =   9
      Top             =   6240
      Width           =   1455
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Previous"
      Height          =   735
      Left            =   4080
      TabIndex        =   8
      Top             =   6240
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Next"
      Height          =   735
      Left            =   5640
      TabIndex        =   7
      Top             =   6240
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Update"
      Height          =   735
      Left            =   5760
      TabIndex        =   6
      Top             =   5280
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Add"
      Height          =   735
      Left            =   2400
      TabIndex        =   5
      Top             =   5280
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Edit"
      Height          =   735
      Left            =   4080
      TabIndex        =   4
      Top             =   5280
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      DataField       =   "desclong"
      DataSource      =   "Data1"
      Height          =   3735
      Left            =   2280
      TabIndex        =   3
      Top             =   1320
      Width           =   6735
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Users\kools\Desktop\college protal\STprjkt.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   9480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Notice"
      Top             =   2400
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "BACK"
      Height          =   735
      Left            =   7320
      TabIndex        =   2
      Top             =   5280
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "NOTICE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   1
      Top             =   480
      Width           =   3615
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      DataField       =   "desclong"
      DataSource      =   "Data1"
      Height          =   3495
      Left            =   2400
      TabIndex        =   0
      Top             =   1440
      Width           =   6495
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim login As Boolean

Private Sub Command1_Click()
Form3.Show
Unload Form10
End Sub

Private Sub Command2_Click()
Text1.Visible = True
Label1.Visible = False
Command3.Visible = False
Data1.Recordset.Edit
End Sub

Private Sub Command3_Click()
Text1.Visible = True
Label1.Visible = False
Command2.Visible = False
Data1.Recordset.AddNew


End Sub

Private Sub Command4_Click()
Command2.Visible = True
Command3.Visible = True
Text1.Visible = False
Label1.Visible = True
Data1.Recordset.Update

Data1.Recordset.MoveLast
End Sub

Private Sub Command5_Click()
If Data1.Recordset.EOF = True Then
Data1.Recordset.MoveFirst
Else
Data1.Recordset.MoveNext
End If
End Sub

Private Sub Command6_Click()
If Data1.Recordset.BOF = True Then
Data1.Recordset.MoveLast
Else
Data1.Recordset.MovePrevious
End If
End Sub

Private Sub Command7_Click()

Data1.Recordset.Delete
Data1.Recordset.MoveLast

End Sub


Private Sub Command8_Click()
If Form3.Label2.Caption = "" Then
    MsgBox "PLZ LOGIN FIRST"
    Form1.Show
    Unload Form10
    Else
    Label1.Visible = True
    Text1.Visible = False
    Command3.Visible = True
    Command2.Visible = True
    Command4.Visible = True
    Command7.Visible = True
End If
End Sub

Private Sub Form_Load()
Data1.Visible = False
Label1.Visible = True
Text1.Visible = False
Command3.Visible = False
Command2.Visible = False
Command4.Visible = False
Command7.Visible = False
End Sub

