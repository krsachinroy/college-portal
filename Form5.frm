VERSION 5.00
Begin VB.Form Form5 
   AutoRedraw      =   -1  'True
   Caption         =   "Admission"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form5"
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
      Height          =   420
      Left            =   10320
      TabIndex        =   20
      Top             =   7320
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
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
      Left            =   11880
      TabIndex        =   19
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      Caption         =   "INFORMATION OF STUDENT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5175
      Left            =   7320
      TabIndex        =   10
      Top             =   1440
      Width           =   5895
      Begin VB.TextBox Text4 
         Height          =   1215
         Left            =   1920
         TabIndex        =   18
         Top             =   3000
         Width           =   3495
      End
      Begin VB.TextBox Text3 
         Height          =   495
         Left            =   1920
         TabIndex        =   17
         Top             =   2280
         Width           =   2535
      End
      Begin VB.TextBox Text2 
         Height          =   495
         Left            =   1920
         TabIndex        =   16
         Top             =   1560
         Width           =   3495
      End
      Begin VB.TextBox Text1 
         Height          =   495
         Left            =   1920
         TabIndex        =   15
         Top             =   840
         Width           =   3495
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "ADDRESS"
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
         Left            =   480
         TabIndex        =   14
         Top             =   3000
         Width           =   1200
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "PHONE"
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
         Left            =   480
         TabIndex        =   13
         Top             =   2280
         Width           =   840
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "E-MAIL"
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
         Left            =   480
         TabIndex        =   12
         Top             =   1560
         Width           =   810
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "NAME"
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
         Left            =   480
         TabIndex        =   11
         Top             =   840
         Width           =   690
      End
   End
   Begin VB.Frame Frame2 
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
      Height          =   3615
      Left            =   2760
      TabIndex        =   4
      Top             =   3960
      Width           =   3735
      Begin VB.OptionButton Option7 
         Caption         =   "ELECTRICAL ENGINEERING"
         Height          =   435
         Left            =   600
         TabIndex        =   9
         Top             =   3000
         Width           =   2535
      End
      Begin VB.OptionButton Option6 
         Caption         =   "MECHNICAL ENGINEERING"
         Height          =   375
         Left            =   600
         TabIndex        =   8
         Top             =   2400
         Width           =   2415
      End
      Begin VB.OptionButton Option5 
         Caption         =   "ELECTRONICS AND COMMUNICATION ENGINEERING"
         Height          =   375
         Left            =   600
         TabIndex        =   7
         Top             =   1800
         Width           =   3015
      End
      Begin VB.OptionButton Option4 
         Caption         =   "INFORMATION TECHNOLOGY"
         Height          =   375
         Left            =   600
         TabIndex        =   6
         Top             =   1200
         Width           =   2655
      End
      Begin VB.OptionButton Option3 
         Caption         =   "COMPUTER SCIENCE AND ENGINEERING"
         Height          =   375
         Left            =   600
         TabIndex        =   5
         Top             =   600
         Width           =   2415
      End
   End
   Begin VB.Frame Frame1 
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
      Height          =   2295
      Left            =   2760
      TabIndex        =   1
      Top             =   1440
      Width           =   2655
      Begin VB.OptionButton Option2 
         Caption         =   "M.TECH"
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
         Left            =   480
         TabIndex        =   3
         Top             =   1440
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "B.TECH"
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
         Left            =   480
         TabIndex        =   2
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Admission"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   690
      Left            =   7320
      TabIndex        =   0
      Top             =   240
      Width           =   2775
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database
Dim rst As Recordset
Dim r As Recordset
Dim s1 As Integer
Private Sub Command1_Click()
Set db = OpenDatabase("C:\Users\kools\Desktop\college protal\college.mdb")
Set rst = db.OpenRecordset("select * from admission")
Set r = db.OpenRecordset("select * from admission")
If Option1.Value = False And Option2.Value = False Then
    MsgBox "FILL ALL INFORMATION CORRECTLY"
    Else
        If Option3.Value = False And Option4.Value = False And Option5.Value = False And Option6.Value = False And Option7.Value = False Then
            MsgBox "FILL ALL INFORMATION CORRECTLY"
            Else
                If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Then
                     MsgBox "FILL ALL INFORMATION CORRECTLY"
                    Else
                        rst.AddNew
                        If Option1.Value = True Then
                           rst.Fields("course") = Option1.Caption
                        End If
                        If Option2.Value = True Then
                           rst.Fields("course") = Option2.Caption
                        End If
                        If Option3.Value = True Then
                           rst.Fields("department") = Option3.Caption
                        End If
                        If Option4.Value = True Then
                           rst.Fields("department") = Option4.Caption
                        End If
                        If Option5.Value = True Then
                           rst.Fields("department") = Option5.Caption
                        End If
                        If Option6.Value = True Then
                           rst.Fields("department") = Option6.Caption
                        End If
                        If Option7.Value = True Then
                           rst.Fields("department") = Option7.Caption
                        End If
                        rst.Fields("name") = Text1
                        rst.Fields("email") = Text2
                        rst.Fields("phone") = Text3.Text
                        rst.Fields("address") = Text4
                        If rst.EOF = True Then
                        rst.Fields("roll") = 1
                        Else
                            r.MoveLast
                            s1 = r.Fields("roll")
                            rst.Fields("roll") = (s1 + 1)
                        End If
                        rst.Update
                        If Option1.Value = True Then
                            MsgBox "Pay the amount of Rs.40000 as demand draft in favour of MCKVIE to college"
                            Else
                                MsgBox "Pay the amount of Rs.80000 as demand draft in favour of MCKVIE to college"
                        End If
                End If
        End If
End If
End Sub

Private Sub Command2_Click()
Form4.Show
Unload Form5
End Sub

