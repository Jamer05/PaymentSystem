VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form BSCPE_BALANCE 
   BackColor       =   &H00004000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BSCPE_BALANCE"
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10695
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   10695
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00008000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   6015
      Index           =   0
      Left            =   720
      TabIndex        =   0
      Top             =   720
      Width           =   9255
      Begin VB.CommandButton Command1 
         Cancel          =   -1  'True
         Caption         =   "Return"
         Default         =   -1  'True
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
         Left            =   8160
         TabIndex        =   20
         Top             =   5640
         Width           =   1095
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H0000C000&
         Height          =   2895
         Left            =   4680
         TabIndex        =   9
         Top             =   2640
         Width           =   3975
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   "Amount"
            ForeColor       =   &H8000000B&
            Height          =   255
            Left            =   1320
            TabIndex        =   19
            Top             =   1800
            Width           =   735
         End
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            DataField       =   "RecAmount"
            DataSource      =   "Adodc1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000B&
            Height          =   375
            Left            =   2040
            TabIndex        =   18
            Top             =   1680
            Width           =   855
         End
         Begin VB.Label Label10 
            DataField       =   "TypeofPayment"
            DataSource      =   "Adodc1"
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
            Left            =   840
            TabIndex        =   17
            Top             =   2160
            Width           =   1455
         End
         Begin VB.Label Label9 
            BackColor       =   &H0000C000&
            Caption         =   "Type"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   240
            TabIndex        =   16
            Top             =   2280
            Width           =   495
         End
         Begin VB.Label Label8 
            BackColor       =   &H0000C000&
            Caption         =   "Time"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   2640
            TabIndex        =   15
            Top             =   720
            Width           =   735
         End
         Begin VB.Label Label7 
            BackColor       =   &H0000C000&
            Caption         =   "Date"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000B&
            Height          =   255
            Left            =   720
            TabIndex        =   14
            Top             =   720
            Width           =   735
         End
         Begin VB.Label Label6 
            BackColor       =   &H0000C000&
            DataField       =   "TimeofPaid"
            DataSource      =   "Adodc1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000B&
            Height          =   495
            Left            =   2160
            TabIndex        =   13
            Top             =   960
            Width           =   1575
         End
         Begin VB.Label Label2 
            BackColor       =   &H0000C000&
            DataField       =   "DateofPaid"
            DataSource      =   "Adodc1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   120
            TabIndex        =   12
            Top             =   960
            Width           =   1695
         End
         Begin VB.Label Label12 
            DataField       =   "MiscellanousType"
            DataSource      =   "Adodc1"
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
            Left            =   2400
            TabIndex        =   11
            Top             =   2160
            Width           =   1455
         End
         Begin VB.Label Label19 
            BackStyle       =   0  'Transparent
            DataField       =   "Sem"
            DataSource      =   "Adodc1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000B&
            Height          =   375
            Left            =   1200
            TabIndex        =   10
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H0000C000&
         Height          =   2895
         Index           =   1
         Left            =   480
         TabIndex        =   4
         Top             =   2640
         Width           =   4215
         Begin VB.Label Label13 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2640
            TabIndex        =   23
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label5 
            Height          =   2775
            Left            =   240
            TabIndex        =   8
            Top             =   120
            Width           =   2175
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "Type"
            ForeColor       =   &H8000000B&
            Height          =   255
            Left            =   3120
            TabIndex        =   7
            Top             =   1320
            Width           =   495
         End
         Begin VB.Label Label15 
            Height          =   375
            Left            =   2640
            TabIndex        =   6
            Top             =   2040
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "Miscellanous Tyoe"
            ForeColor       =   &H8000000B&
            Height          =   255
            Left            =   2640
            TabIndex        =   5
            Top             =   2520
            Visible         =   0   'False
            Width           =   1455
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H0000C000&
         Caption         =   "Balance"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   1095
         Left            =   1440
         TabIndex        =   1
         Top             =   360
         Width           =   6495
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            DataField       =   "Balance"
            DataSource      =   "Adodc1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   39
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   855
            Left            =   3120
            TabIndex        =   3
            Top             =   120
            Width           =   2535
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "PHP."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   39
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   975
            Left            =   960
            TabIndex        =   2
            Top             =   120
            Width           =   2535
         End
      End
      Begin VB.Label Label11 
         BackColor       =   &H00008000&
         Caption         =   "Recent:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   555
         Left            =   480
         TabIndex        =   22
         Top             =   2160
         Width           =   1845
      End
      Begin VB.Label Label4 
         BackColor       =   &H00008000&
         Caption         =   "History:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   4680
         TabIndex        =   21
         Top             =   2160
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   11520
      Top             =   7440
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"BSCPE_BALANCE.frx":0000
      OLEDBString     =   $"BSCPE_BALANCE.frx":0099
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Table2"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "BSCPE_BALANCE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
BSCPE_BALANCE.Hide
BSCPE_ACCOUNT.Show

End Sub

Private Sub Command1_KeyPress(KeyAscii As Integer)
If Chr(KeyAscii) = Chr(27) Then
Call Command1_Click
End If
End Sub

Private Sub Command2_Click()
Label13.Caption = "jd"
End Sub
