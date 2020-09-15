VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form BSIT 
   BackColor       =   &H00008000&
   Caption         =   "Form1"
   ClientHeight    =   6120
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9960
   LinkTopic       =   "Form1"
   ScaleHeight     =   6120
   ScaleWidth      =   9960
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   960
      Top             =   5280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0000C000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3975
      Left            =   600
      TabIndex        =   0
      Top             =   1080
      Width           =   8775
      Begin VB.CommandButton Command5 
         Caption         =   "Back"
         Height          =   255
         Left            =   7800
         TabIndex        =   17
         Top             =   3480
         Width           =   735
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2520
         TabIndex        =   5
         Top             =   2280
         Width           =   2655
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Add"
         Default         =   -1  'True
         Height          =   375
         Left            =   2520
         TabIndex        =   4
         Top             =   3120
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   4080
         TabIndex        =   3
         Top             =   3120
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Print"
         Height          =   255
         Left            =   6240
         TabIndex        =   2
         Top             =   3480
         Width           =   855
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Clear"
         Height          =   255
         Left            =   7080
         TabIndex        =   1
         Top             =   3480
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
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
         Left            =   600
         TabIndex        =   15
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label2 
         DataField       =   "Uname"
         DataSource      =   "Adodc2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2520
         TabIndex        =   14
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Balance:"
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
         Height          =   375
         Left            =   600
         TabIndex        =   13
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Add Amount"
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
         Height          =   375
         Left            =   600
         TabIndex        =   12
         Top             =   2400
         Width           =   2055
      End
      Begin VB.Label Label4 
         DataField       =   "Balance"
         DataSource      =   "Adodc2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2520
         TabIndex        =   11
         Top             =   1320
         Width           =   2655
      End
      Begin VB.Label Label7 
         Height          =   2295
         Left            =   6360
         TabIndex        =   10
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Label8 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6840
         TabIndex        =   9
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
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
         Left            =   5520
         TabIndex        =   8
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Label10 
         DataField       =   "Balance"
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
         Height          =   375
         Left            =   7200
         TabIndex        =   7
         Top             =   3000
         Width           =   1335
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "New Balance"
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
         Left            =   5520
         TabIndex        =   6
         Top             =   3120
         Width           =   1575
      End
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   4320
      Top             =   5520
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
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
      Connect         =   $"BSIT.frx":0000
      OLEDBString     =   $"BSIT.frx":0099
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Table1"
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
      Height          =   330
      Left            =   4320
      Top             =   5160
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
      Connect         =   $"BSIT.frx":0132
      OLEDBString     =   $"BSIT.frx":01CB
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from Table1"
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
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "STUDENTS INFORMATION"
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
      Height          =   615
      Left            =   1680
      TabIndex        =   16
      Top             =   0
      Width           =   6855
   End
End
Attribute VB_Name = "BSIT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim Add As Double
Dim Amount As Double
Adodc1.RecordSource = "select * from Table1 where Balance"
Adodc1.Refresh
Amount = Val(Text2.Text)
Balance = Adodc1.Recordset![Balance]
Add = Balance + Amount
    If Text2.Text = "" Then
    MsgBox "Please Enter Amount", vbCritical
    Label10.Caption = ""
    ElseIf Text2.Text = 0 Then
    MsgBox "Invalid Amount", vbCritical
Else
 Adodc1.Recordset.Update
 Label8.Caption = Amount
Label10.Caption = Add
Label7.Caption = "***********************************" + vbNewLine + Space(13) & "Transaction" + vbNewLine + "***********************************" + vbNewLine + "Time:" + vbNewLine + FormatDateTime(Now, vbShortTime) + vbNewLine + "Date:" + vbNewLine + FormatDateTime(Now, vbShortDate) + vbNewLine + "Current Balance:" & Label4.Caption + vbNewLine + "Amount:" & Text2.Text + vbNewLine + "***************************************" + vbNewLine + "New Balance" & Label10.Caption
 Text2.Text = ""
 Text2.SetFocus
 End If
End Sub

Private Sub Command2_Click()
BSIT.Hide
Deposit.Show
Deposit.Text1.Text = ""
Deposit.Text1.SetFocus
End Sub

Private Sub Command3_Click()
Adodc1.Recordset.AddNew
Label10.Caption = Add

If Label8.Caption = "" Then
MsgBox "Unable to Save", vbCritical

Else
Label7.Caption = ""
Label10.Caption = Add
Text2.SetFocus
Label8.Caption = ""
CommonDialog1.ShowPrinter

End If

End Sub

Private Sub Command4_Click()
Text2.Text = ""
Label8.Caption = ""
Label7.Caption = ""
Label10.Caption = ""
End Sub

Private Sub Command5_Click()
BSIT.Hide
Deposit.Show
Deposit.Text1.Text = ""
Deposit.Text1.SetFocus
End Sub

Private Sub Form_Load()
Label10.Caption = ""
End Sub

