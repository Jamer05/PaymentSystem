VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form BSBA_LOGIN 
   BackColor       =   &H00004000&
   Caption         =   "Form3"
   ClientHeight    =   6120
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8685
   HasDC           =   0   'False
   LinkTopic       =   "Form3"
   ScaleHeight     =   6120
   ScaleWidth      =   8685
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   6705
      Left            =   8760
      Top             =   6240
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   11827
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
      Connect         =   $"BSBA_LOGIN.frx":0000
      OLEDBString     =   $"BSBA_LOGIN.frx":0099
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from Table3"
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
   Begin VB.TextBox Text2 
      DataField       =   "Pass"
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
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   4080
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   4080
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      DataField       =   "User"
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
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   4080
      TabIndex        =   3
      Top             =   3480
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   1
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00004000&
      Caption         =   "ShowPassword"
      ForeColor       =   &H8000000B&
      Height          =   195
      Left            =   4560
      TabIndex        =   0
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "PassWord"
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
      Height          =   375
      Left            =   2280
      TabIndex        =   8
      Top             =   4080
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackColor       =   &H00004000&
      Caption         =   "UserName"
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
      Height          =   375
      Left            =   2280
      TabIndex        =   7
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Log In"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   48.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   3240
      TabIndex        =   6
      Top             =   1200
      Width           =   3015
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "BSBA"
      BeginProperty Font 
         Name            =   "Roman"
         Size            =   24
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   615
      Left            =   4200
      TabIndex        =   5
      Top             =   2640
      Width           =   1200
   End
End
Attribute VB_Name = "BSBA_LOGIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
If Check1 = 0 Then
     Text2.PasswordChar = "*"

ElseIf Check1 = 1 Then
    Text2.PasswordChar = Char

End If
End Sub

Private Sub Command1_Click()
Text2.SetFocus
Adodc1.RecordSource = "select* from Table3 where User ='" + Text1.Text + "' and Pass='" + Text2.Text + "'"
Adodc1.Refresh
Text1.SetFocus
If Adodc1.Recordset.EOF Then
Text1.SetFocus
Text1.Text = ""
Text2.Text = ""
MsgBox "invalid username/password", vbCritical
Text2.SetFocus
Else

MsgBox " Successfully Login !!", vbInformation
BSBA_LOGIN.Hide
BSBA_ACCOUNT.Show
End If
End Sub

Private Sub Command2_Click()
BSBA_LOGIN.Hide
Course.Show

End Sub

Private Sub Form_Load()
Text1.Text = ""
Text2.Text = ""
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If Chr(KeyAscii) = Chr(13) Then
Text2.SetFocus
ElseIf Chr(KeyAscii) = Chr(27) Then
Call Command1_Click
Course.Combo1.Text = ""
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)

If Chr(KeyAscii) = Chr(13) Then
Call Command1_Click
ElseIf Chr(KeyAscii) = Chr(27) Then
Call Command1_Click
Course.Combo1.Text = ""
End If
End Sub
