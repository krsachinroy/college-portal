VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form11 
   Caption         =   "LIBRARY DISPLAY"
   ClientHeight    =   9660
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16485
   LinkTopic       =   "Form3"
   ScaleHeight     =   9660
   ScaleWidth      =   16485
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command4 
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
      Left            =   14880
      TabIndex        =   73
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   255
      Left            =   480
      TabIndex        =   72
      Top             =   1680
      Width           =   495
   End
   Begin VB.ComboBox Combo8 
      Height          =   315
      Left            =   5880
      TabIndex        =   46
      Top             =   360
      Width           =   1335
   End
   Begin VB.ComboBox Combo7 
      Height          =   315
      Left            =   3000
      TabIndex        =   44
      Top             =   360
      Width           =   975
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8415
      Left            =   1680
      TabIndex        =   0
      Top             =   960
      Width           =   12855
      _ExtentX        =   22675
      _ExtentY        =   14843
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "ISSUE BOOK"
      TabPicture(0)   =   "Form11.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Command1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1(2)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Data1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Data2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "SUBMIT BOOK"
      TabPicture(1)   =   "Form11.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1(1)"
      Tab(1).Control(1)=   "Command2"
      Tab(1).ControlCount=   2
      Begin VB.Data Data2 
         Caption         =   "Data2"
         Connect         =   "Access"
         DatabaseName    =   "C:\Users\kools\Desktop\college protal\database2.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   1215
         Left            =   10080
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "member detail"
         Top             =   1920
         Width           =   2535
      End
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   "C:\Users\kools\Desktop\college protal\database2.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   1095
         Left            =   10200
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "book detail"
         Top             =   4920
         Width           =   2175
      End
      Begin VB.Frame Frame1 
         Caption         =   "MEMBER DETAIL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6975
         Index           =   2
         Left            =   1800
         TabIndex        =   47
         Top             =   720
         Width           =   8175
         Begin MSACAL.Calendar Calendar1 
            Height          =   2175
            Left            =   2880
            TabIndex        =   66
            Top             =   5040
            Width           =   4095
            _Version        =   524288
            _ExtentX        =   7223
            _ExtentY        =   3836
            _StockProps     =   1
            BackColor       =   -2147483633
            Year            =   2019
            Month           =   4
            Day             =   8
            DayLength       =   1
            MonthLength     =   2
            DayFontColor    =   0
            FirstDay        =   1
            GridCellEffect  =   1
            GridFontColor   =   10485760
            GridLinesColor  =   -2147483632
            ShowDateSelectors=   -1  'True
            ShowDays        =   -1  'True
            ShowHorizontalGrid=   -1  'True
            ShowTitle       =   -1  'True
            ShowVerticalGrid=   -1  'True
            TitleFontColor  =   10485760
            ValueIsNull     =   0   'False
            BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.TextBox Text2 
            DataField       =   "member name"
            DataSource      =   "Data2"
            Height          =   375
            Index           =   2
            Left            =   2880
            TabIndex        =   60
            Top             =   1200
            Width           =   1335
         End
         Begin VB.Frame Frame2 
            Caption         =   "BOOK DETAIL"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4695
            Index           =   2
            Left            =   120
            TabIndex        =   49
            Top             =   2280
            Width           =   7695
            Begin MSACAL.Calendar Calendar2 
               Height          =   855
               Left            =   2760
               TabIndex        =   67
               Top             =   3840
               Width           =   2415
               _Version        =   524288
               _ExtentX        =   4260
               _ExtentY        =   1508
               _StockProps     =   1
               BackColor       =   -2147483633
               Year            =   2019
               Month           =   4
               Day             =   8
               DayLength       =   1
               MonthLength     =   2
               DayFontColor    =   0
               FirstDay        =   1
               GridCellEffect  =   1
               GridFontColor   =   10485760
               GridLinesColor  =   -2147483632
               ShowDateSelectors=   -1  'True
               ShowDays        =   -1  'True
               ShowHorizontalGrid=   -1  'True
               ShowTitle       =   -1  'True
               ShowVerticalGrid=   -1  'True
               TitleFontColor  =   10485760
               ValueIsNull     =   0   'False
               BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin VB.TextBox Text8 
               DataField       =   "author "
               DataSource      =   "Data1"
               Height          =   285
               Left            =   2640
               TabIndex        =   65
               Top             =   1680
               Width           =   3255
            End
            Begin VB.TextBox Text6 
               DataField       =   "avialable stock"
               DataSource      =   "Data1"
               Height          =   495
               Index           =   2
               Left            =   6000
               TabIndex        =   53
               Top             =   2280
               Width           =   1575
            End
            Begin VB.TextBox Text5 
               DataField       =   "total stock"
               DataSource      =   "Data1"
               Height          =   285
               Index           =   2
               Left            =   2280
               TabIndex        =   52
               Top             =   2400
               Width           =   975
            End
            Begin VB.TextBox Text4 
               DataField       =   "book title"
               DataSource      =   "Data1"
               Height          =   285
               Index           =   2
               Left            =   2640
               TabIndex        =   51
               Top             =   1080
               Width           =   3855
            End
            Begin VB.TextBox Text3 
               DataField       =   "code"
               DataSource      =   "Data1"
               Height          =   375
               Index           =   2
               Left            =   2640
               TabIndex        =   50
               Top             =   480
               Width           =   735
            End
            Begin VB.Label Label15 
               Caption         =   "AUTHOR  :"
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
               Left            =   360
               TabIndex        =   69
               Top             =   1680
               Width           =   1335
            End
            Begin VB.Label Label8 
               Caption         =   "LAST SUBMIT DATE:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   615
               Index           =   2
               Left            =   720
               TabIndex        =   59
               Top             =   3840
               Width           =   1575
            End
            Begin VB.Label Label7 
               Caption         =   "ISSUE DATE:"
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
               Index           =   2
               Left            =   600
               TabIndex        =   58
               Top             =   3000
               Width           =   1575
            End
            Begin VB.Label Label6 
               Caption         =   "AVAILABLE STOCK:"
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
               Index           =   2
               Left            =   3480
               TabIndex        =   57
               Top             =   2400
               Width           =   2175
            End
            Begin VB.Label Label5 
               Caption         =   "TOTAL STOCK:"
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
               Index           =   2
               Left            =   360
               TabIndex        =   56
               Top             =   2400
               Width           =   1815
            End
            Begin VB.Label Label4 
               Caption         =   "BOOK TITLE :"
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
               Index           =   2
               Left            =   360
               TabIndex        =   55
               Top             =   1080
               Width           =   1575
            End
            Begin VB.Label Label3 
               Caption         =   "CODE:"
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
               Index           =   2
               Left            =   360
               TabIndex        =   54
               Top             =   480
               Width           =   1095
            End
         End
         Begin VB.TextBox Text1 
            DataField       =   "code"
            DataSource      =   "Data2"
            Height          =   285
            Index           =   2
            Left            =   2880
            TabIndex        =   48
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label Label2 
            Caption         =   "MEMBER NAME :"
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
            Index           =   2
            Left            =   480
            TabIndex        =   62
            Top             =   1200
            Width           =   1935
         End
         Begin VB.Label Label1 
            Caption         =   "CODE :"
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
            Index           =   2
            Left            =   480
            TabIndex        =   61
            Top             =   480
            Width           =   1455
         End
      End
      Begin VB.CommandButton Command2 
         Caption         =   "SUBMIT"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   -66360
         TabIndex        =   42
         Top             =   7680
         Width           =   1695
      End
      Begin VB.Frame Frame1 
         Caption         =   "MEMBER DETAIL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6975
         Index           =   1
         Left            =   -72960
         TabIndex        =   26
         Top             =   600
         Width           =   8175
         Begin VB.TextBox Text1 
            DataField       =   "code"
            DataSource      =   "Data2"
            Height          =   285
            Index           =   1
            Left            =   2280
            TabIndex        =   39
            Top             =   480
            Width           =   1215
         End
         Begin VB.Frame Frame2 
            Caption         =   "BOOK DETAIL"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4695
            Index           =   1
            Left            =   240
            TabIndex        =   28
            Top             =   2160
            Width           =   7695
            Begin MSACAL.Calendar Calendar4 
               Height          =   735
               Left            =   2760
               TabIndex        =   71
               Top             =   3720
               Width           =   2295
               _Version        =   524288
               _ExtentX        =   4048
               _ExtentY        =   1296
               _StockProps     =   1
               BackColor       =   -2147483633
               Year            =   2019
               Month           =   4
               Day             =   8
               DayLength       =   1
               MonthLength     =   2
               DayFontColor    =   0
               FirstDay        =   1
               GridCellEffect  =   1
               GridFontColor   =   10485760
               GridLinesColor  =   -2147483632
               ShowDateSelectors=   -1  'True
               ShowDays        =   -1  'True
               ShowHorizontalGrid=   -1  'True
               ShowTitle       =   -1  'True
               ShowVerticalGrid=   -1  'True
               TitleFontColor  =   10485760
               ValueIsNull     =   0   'False
               BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin MSACAL.Calendar Calendar3 
               Height          =   615
               Left            =   2160
               TabIndex        =   70
               Top             =   2880
               Width           =   2415
               _Version        =   524288
               _ExtentX        =   4260
               _ExtentY        =   1085
               _StockProps     =   1
               BackColor       =   -2147483633
               Year            =   2019
               Month           =   4
               Day             =   8
               DayLength       =   1
               MonthLength     =   2
               DayFontColor    =   0
               FirstDay        =   1
               GridCellEffect  =   1
               GridFontColor   =   10485760
               GridLinesColor  =   -2147483632
               ShowDateSelectors=   -1  'True
               ShowDays        =   -1  'True
               ShowHorizontalGrid=   -1  'True
               ShowTitle       =   -1  'True
               ShowVerticalGrid=   -1  'True
               TitleFontColor  =   10485760
               ValueIsNull     =   0   'False
               BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin VB.TextBox Text7 
               DataField       =   "author "
               DataSource      =   "Data1"
               Height          =   375
               Left            =   2400
               TabIndex        =   64
               Top             =   1440
               Width           =   4095
            End
            Begin VB.TextBox Text3 
               DataField       =   "code"
               DataSource      =   "Data1"
               Height          =   375
               Index           =   1
               Left            =   2400
               TabIndex        =   32
               Top             =   480
               Width           =   735
            End
            Begin VB.TextBox Text4 
               DataField       =   "book title"
               DataSource      =   "Data1"
               Height          =   285
               Index           =   1
               Left            =   2400
               TabIndex        =   31
               Top             =   1080
               Width           =   3855
            End
            Begin VB.TextBox Text5 
               DataField       =   "total stock"
               DataSource      =   "Data1"
               Height          =   285
               Index           =   1
               Left            =   1800
               TabIndex        =   30
               Top             =   2400
               Width           =   975
            End
            Begin VB.TextBox Text6 
               DataField       =   "avialable stock"
               DataSource      =   "Data1"
               Height          =   495
               Index           =   1
               Left            =   4560
               TabIndex        =   29
               Top             =   2160
               Width           =   1575
            End
            Begin VB.Label Label13 
               Caption         =   "AUTHOR NAME:"
               DataSource      =   "Data1"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   360
               TabIndex        =   63
               Top             =   1560
               Width           =   1455
            End
            Begin VB.Label Label3 
               Caption         =   "CODE"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   1
               Left            =   480
               TabIndex        =   38
               Top             =   480
               Width           =   1095
            End
            Begin VB.Label Label4 
               Caption         =   "BOOK TITLE"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   1
               Left            =   480
               TabIndex        =   37
               Top             =   1200
               Width           =   1335
            End
            Begin VB.Label Label5 
               Caption         =   "TOTAL STOCK"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   1
               Left            =   600
               TabIndex        =   36
               Top             =   2280
               Width           =   855
            End
            Begin VB.Label Label6 
               Caption         =   "AVAILABLE STOCK"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   1
               Left            =   3240
               TabIndex        =   35
               Top             =   2280
               Width           =   1095
            End
            Begin VB.Label Label7 
               Caption         =   "ISSUE DATE"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Index           =   1
               Left            =   720
               TabIndex        =   34
               Top             =   2880
               Width           =   1575
            End
            Begin VB.Label Label8 
               Caption         =   "LAST SUBMIT DATE"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   615
               Index           =   1
               Left            =   720
               TabIndex        =   33
               Top             =   3720
               Width           =   1575
            End
         End
         Begin VB.TextBox Text2 
            DataField       =   "member name"
            DataSource      =   "Data2"
            Height          =   375
            Index           =   1
            Left            =   2400
            TabIndex        =   27
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "CODE :"
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
            Index           =   1
            Left            =   480
            TabIndex        =   41
            Top             =   480
            Width           =   1455
         End
         Begin VB.Label Label2 
            Caption         =   "MEMBER NAME :"
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
            Index           =   1
            Left            =   480
            TabIndex        =   40
            Top             =   1080
            Width           =   1815
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "ISSUE"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5760
         TabIndex        =   25
         Top             =   7800
         Width           =   1815
      End
      Begin VB.Frame Frame1 
         Caption         =   "member detail"
         Height          =   6855
         Index           =   0
         Left            =   -74400
         TabIndex        =   1
         Top             =   600
         Width           =   7935
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   0
            Left            =   1560
            TabIndex        =   22
            Top             =   480
            Width           =   1215
         End
         Begin VB.Frame Frame2 
            Caption         =   "BOOK DETAIL"
            Height          =   4575
            Index           =   0
            Left            =   240
            TabIndex        =   3
            Top             =   1800
            Width           =   7215
            Begin VB.TextBox Text3 
               Height          =   375
               Index           =   0
               Left            =   1560
               TabIndex        =   13
               Top             =   480
               Width           =   735
            End
            Begin VB.TextBox Text4 
               Height          =   285
               Index           =   0
               Left            =   1920
               TabIndex        =   12
               Top             =   1080
               Width           =   3855
            End
            Begin VB.TextBox Text5 
               Height          =   285
               Index           =   0
               Left            =   1800
               TabIndex        =   11
               Top             =   1920
               Width           =   975
            End
            Begin VB.TextBox Text6 
               Height          =   495
               Index           =   0
               Left            =   4440
               TabIndex        =   10
               Top             =   1920
               Width           =   1575
            End
            Begin VB.ComboBox Combo1 
               Height          =   315
               Index           =   0
               Left            =   2160
               TabIndex        =   9
               Text            =   "Combo1"
               Top             =   2880
               Width           =   735
            End
            Begin VB.ComboBox Combo2 
               Height          =   315
               Index           =   0
               Left            =   2880
               TabIndex        =   8
               Text            =   "Combo2"
               Top             =   2880
               Width           =   735
            End
            Begin VB.ComboBox Combo3 
               Height          =   315
               Index           =   0
               Left            =   3600
               TabIndex        =   7
               Text            =   "Combo3"
               Top             =   2880
               Width           =   615
            End
            Begin VB.ComboBox Combo4 
               Height          =   315
               Index           =   0
               Left            =   2760
               TabIndex        =   6
               Text            =   "Combo4"
               Top             =   3720
               Width           =   735
            End
            Begin VB.ComboBox Combo5 
               Height          =   315
               Index           =   0
               Left            =   3480
               TabIndex        =   5
               Text            =   "Combo5"
               Top             =   3720
               Width           =   615
            End
            Begin VB.ComboBox Combo6 
               Height          =   315
               Index           =   0
               Left            =   4080
               TabIndex        =   4
               Text            =   "Combo6"
               Top             =   3720
               Width           =   615
            End
            Begin VB.Label Label3 
               Caption         =   "CODE"
               Height          =   375
               Index           =   0
               Left            =   480
               TabIndex        =   21
               Top             =   480
               Width           =   1095
            End
            Begin VB.Label Label4 
               Caption         =   "BOOK TITLE"
               Height          =   375
               Index           =   0
               Left            =   480
               TabIndex        =   20
               Top             =   1080
               Width           =   1095
            End
            Begin VB.Label Label5 
               Caption         =   "TOTAL STOCK"
               Height          =   375
               Index           =   0
               Left            =   600
               TabIndex        =   19
               Top             =   1920
               Width           =   855
            End
            Begin VB.Label Label6 
               Caption         =   "AVAILABLE STOCK"
               Height          =   735
               Index           =   0
               Left            =   3240
               TabIndex        =   18
               Top             =   1920
               Width           =   1095
            End
            Begin VB.Label Label7 
               Caption         =   "ISSUE DATE"
               Height          =   495
               Index           =   0
               Left            =   720
               TabIndex        =   17
               Top             =   2880
               Width           =   1575
            End
            Begin VB.Label Label8 
               Caption         =   "LAST SUBMIT DATE"
               Height          =   615
               Index           =   0
               Left            =   720
               TabIndex        =   16
               Top             =   3720
               Width           =   1575
            End
            Begin VB.Label Label9 
               Caption         =   "DD-MM-YYYY"
               Height          =   255
               Index           =   0
               Left            =   4320
               TabIndex        =   15
               Top             =   2880
               Width           =   1575
            End
            Begin VB.Label Label10 
               Caption         =   "DD-MM-YYYY"
               Height          =   375
               Index           =   0
               Left            =   4920
               TabIndex        =   14
               Top             =   3720
               Width           =   1455
            End
         End
         Begin VB.TextBox Text2 
            Height          =   375
            Index           =   0
            Left            =   1560
            TabIndex        =   2
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "CODE :"
            Height          =   375
            Index           =   0
            Left            =   480
            TabIndex        =   24
            Top             =   480
            Width           =   1455
         End
         Begin VB.Label Label2 
            Caption         =   "MEMBER NAME :"
            Height          =   375
            Index           =   0
            Left            =   480
            TabIndex        =   23
            Top             =   1080
            Width           =   735
         End
      End
   End
   Begin VB.Label Label14 
      Caption         =   "Label14"
      Height          =   495
      Left            =   7680
      TabIndex        =   68
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label12 
      Caption         =   "YEAR"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      TabIndex        =   45
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label11 
      Caption         =   "CLASS"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   43
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Calendar1_Click()
MsgBox Calendar1.Value
End Sub

Private Sub Command1_Click()

Dim class As String
Dim year As String
class = Combo7.Text
year = Combo8.Text

If class = "" And year = "" Then

MsgBox "sorry book will not be issued"
Else
MsgBox "book issued"
Form4.Show
End If
End Sub

Private Sub Command2_Click()
Dim class As String
Dim year As String
class = Combo7.Text
year = Combo8.Text

If class = "" And year = "" Then

MsgBox "sorry book will not be sumitted"
Else
MsgBox "book will be submitted"
End If
End Sub

Private Sub Command4_Click()
Form3.Show
Unload Form11
End Sub

Private Sub Form_Load()
Combo7.AddItem "BBA"
Combo7.AddItem "B.TECH"
Combo7.AddItem " MCA"
Combo8.AddItem "FY"
Combo8.AddItem "3RD YEAR"
Combo8.AddItem "2ND YEAR"
Combo8.AddItem "1ST YEAR"
End Sub

