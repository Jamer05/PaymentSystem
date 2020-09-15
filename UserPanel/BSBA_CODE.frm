VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form BSBA_CODE 
   BackColor       =   &H00004000&
   Caption         =   "BSBA_CODE"
   ClientHeight    =   3120
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4725
   LinkTopic       =   "Form2"
   ScaleHeight     =   3120
   ScaleWidth      =   4725
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   4320
      Top             =   3360
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
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
      Connect         =   $"BSBA_CODE.frx":0000
      OLEDBString     =   $"BSBA_CODE.frx":0099
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
   Begin VB.Frame Frame1 
      BackColor       =   &H0000C000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2415
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   4095
      Begin VB.TextBox Text1 
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
         IMEMode         =   3  'DISABLE
         Left            =   960
         MaxLength       =   8
         PasswordChar    =   "X"
         TabIndex        =   2
         Top             =   1080
         Width           =   2295
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Continue"
         Height          =   375
         Left            =   1560
         TabIndex        =   1
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Please Enter Security Code"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   495
         Left            =   360
         TabIndex        =   3
         Top             =   360
         Width           =   3495
      End
   End
End
Attribute VB_Name = "BSBA_CODE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text1.SetFocus
Adodc1.RecordSource = "select * from Table3 where PinCode ='" + Text1.Text + "'"
Adodc1.Refresh
Text1.SetFocus
If Adodc1.Recordset.EOF Then
Text1.SetFocus
Text1.Text = ""

MsgBox "invalid Code", vbCritical
Else

MsgBox " Successfully Paid !!", vbInformation
BSBA_CODE.Hide
BSBA_ACCOUNT.Show
BSBA_PAY.Hide

End If
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
