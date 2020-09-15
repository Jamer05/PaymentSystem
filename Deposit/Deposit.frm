VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Deposit 
   BackColor       =   &H00004000&
   Caption         =   "Form1"
   ClientHeight    =   7440
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13005
   LinkTopic       =   "Form1"
   ScaleHeight     =   7440
   ScaleWidth      =   13005
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   495
      Left            =   8400
      Top             =   8880
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
      CommandType     =   1
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
      Connect         =   $"Deposit.frx":0000
      OLEDBString     =   $"Deposit.frx":0099
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from Table3"
      Caption         =   "Adodc3"
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   9600
      Top             =   8640
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
      Connect         =   $"Deposit.frx":0132
      OLEDBString     =   $"Deposit.frx":01CB
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from Table2"
      Caption         =   "Adodc2"
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   11400
      Top             =   9000
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
      Connect         =   $"Deposit.frx":0264
      OLEDBString     =   $"Deposit.frx":02FD
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select*from Table1"
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
   Begin VB.Frame Frame3 
      BackColor       =   &H0000C000&
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   4815
      Left            =   9000
      TabIndex        =   2
      Top             =   1920
      Width           =   3495
      Begin VB.CommandButton Command3 
         Caption         =   "Check"
         CausesValidation=   0   'False
         Height          =   495
         Left            =   1200
         TabIndex        =   15
         Top             =   3600
         Width           =   1335
      End
      Begin VB.TextBox Text3 
         DataField       =   "RefNumber"
         DataSource      =   "Adodc3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   720
         MaxLength       =   10
         TabIndex        =   12
         Top             =   2400
         Width           =   2175
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Referenece No."
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
         Height          =   495
         Index           =   2
         Left            =   240
         TabIndex        =   9
         Top             =   1680
         Width           =   3135
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "BSBA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   735
         Index           =   2
         Left            =   1080
         TabIndex        =   5
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0000C000&
      BorderStyle     =   0  'None
      Caption         =   "CheckCheck"
      Height          =   4815
      Left            =   4800
      TabIndex        =   1
      Top             =   1920
      Width           =   3615
      Begin VB.CommandButton Command2 
         Caption         =   "Check"
         Height          =   495
         Left            =   1200
         TabIndex        =   14
         Top             =   3600
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         DataField       =   "RefNumber"
         DataSource      =   "Adodc2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   720
         MaxLength       =   10
         TabIndex        =   11
         Top             =   2400
         Width           =   2295
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Referenece No."
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
         Height          =   495
         Index           =   1
         Left            =   360
         TabIndex        =   8
         Top             =   1680
         Width           =   3135
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "BSCPE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   735
         Index           =   1
         Left            =   960
         TabIndex        =   4
         Top             =   480
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0000C000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4815
      Left            =   360
      TabIndex        =   0
      Top             =   1920
      Width           =   3855
      Begin VB.CommandButton Command1 
         Caption         =   "Check"
         Height          =   435
         Left            =   1200
         TabIndex        =   13
         Top             =   3720
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         DataField       =   "RefNUmber"
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
         Height          =   495
         Left            =   720
         MaxLength       =   10
         TabIndex        =   10
         Top             =   2400
         Width           =   2295
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Referenece No."
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
         Height          =   495
         Index           =   0
         Left            =   480
         TabIndex        =   7
         Top             =   1680
         Width           =   3135
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "BSIT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   735
         Index           =   0
         Left            =   1320
         TabIndex        =   3
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Deposit Panel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   48.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   1215
      Left            =   3600
      TabIndex        =   6
      Top             =   480
      Width           =   6615
   End
End
Attribute VB_Name = "Deposit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text1.SetFocus
Adodc1.RecordSource = "select * from Table1 where RefNumber ='" + Text1.Text + "'"
Adodc1.Refresh
Text1.SetFocus
If Adodc1.Recordset.EOF Then
Text1.Text = ""
Text1.SetFocus
MsgBox "invalid Reference Number", vbCritical
Else
Text1.SetFocus
BSIT.Text2.Text = ""
Deposit.Hide
BSIT.Show
BSIT.Label7.Caption = ""
BSIT.Label8.Caption = ""
BSIT.Text2.SetFocus
End If
End Sub

Private Sub Command2_Click()
Text2.SetFocus
Adodc2.RecordSource = "select * from Table2 where RefNumber ='" + Text2.Text + "'"
Adodc2.Refresh
Text2.SetFocus
If Adodc2.Recordset.EOF Then
Text2.SetFocus
Text2.Text = ""
MsgBox "invalid Reference Number", vbCritical
Else
BSCPE.Label10.Caption = ""
Text1.SetFocus
BSCPE.Text2.Text = ""
Deposit.Hide
BSCPE.Show
BSCPE.Label8.Caption = ""
BSCPE.Text2.SetFocus
End If
End Sub

Private Sub Command3_Click()
Text3.SetFocus
Adodc3.RecordSource = "select * from Table3 where RefNumber ='" + Text3.Text + "'"
Adodc3.Refresh
Text1.SetFocus
If Adodc3.Recordset.EOF Then
Text2.SetFocus
Text2.Text = ""

MsgBox "invalid Reference Number", vbCritical
Else
BSBA.Label10.Caption = ""
Text1.SetFocus
BSBA.Text2.Text = ""
Deposit.Hide
BSBA.Show
BSBA.Label8.Caption = ""
BSBA.Text2.SetFocus
End If
End Sub

Private Sub Command3_KeyPress(KeyAscii As Integer)
If Chr(KeyAscii) = Chr(13) Then
Call Command3_Click
End If
End Sub

Private Sub Form_Load()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
Dim KeyChar As String
If KeyAscii > 31 Then
KeyChar = Chr(KeyAscii)
If Not IsNumeric(KeyChar) Then
KeyAscii = 0
End If
End If

If Chr(KeyAscii) = Chr(13) Then
Call Command1_Click
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
Dim KeyChar As String
If KeyAscii > 31 Then
KeyChar = Chr(KeyAscii)
If Not IsNumeric(KeyChar) Then
KeyAscii = 0
End If
End If
If Chr(KeyAscii) = Chr(13) Then
Call Command2_Click
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
Dim KeyChar As String
If KeyAscii > 31 Then
KeyChar = Chr(KeyAscii)
If Not IsNumeric(KeyChar) Then
KeyAscii = 0
End If
End If
If Chr(KeyAscii) = Chr(13) Then
Call Command3_Click
End If

End Sub
