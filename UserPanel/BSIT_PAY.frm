VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form BSIT_PAY 
   BackColor       =   &H00004000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " "
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9090
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   9090
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Home"
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
      Left            =   7320
      TabIndex        =   8
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Caption         =   "pay"
      Height          =   2895
      Left            =   720
      TabIndex        =   2
      Top             =   960
      Width           =   4815
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "BSIT_PAY.frx":0000
         Left            =   2760
         List            =   "BSIT_PAY.frx":0010
         TabIndex        =   20
         Text            =   "Choose"
         Top             =   1920
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00008000&
         Caption         =   "Sem"
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   720
         TabIndex        =   16
         Top             =   360
         Width           =   3255
         Begin VB.OptionButton Option2 
            BackColor       =   &H00008000&
            Caption         =   "Second Sem"
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   1800
            TabIndex        =   18
            Top             =   240
            Width           =   1335
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00008000&
            Caption         =   "First Sem"
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   120
            TabIndex        =   17
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Done"
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
         Height          =   615
         Left            =   3720
         TabIndex        =   6
         Top             =   2280
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1920
         MaxLength       =   5
         TabIndex        =   5
         Top             =   2400
         Width           =   1455
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00008000&
         Caption         =   "Type"
         ForeColor       =   &H8000000B&
         Height          =   615
         Left            =   720
         TabIndex        =   3
         Top             =   1320
         Width           =   3255
         Begin VB.CheckBox Check1 
            BackColor       =   &H00008000&
            Caption         =   "Miscellanous"
            ForeColor       =   &H8000000B&
            Height          =   255
            Left            =   1800
            TabIndex        =   19
            Top             =   240
            Width           =   1335
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H00008000&
            Caption         =   "Tuition"
            ForeColor       =   &H8000000B&
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Label Label6 
         BackColor       =   &H00008000&
         Caption         =   "Amount"
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
         Left            =   480
         TabIndex        =   7
         Top             =   2400
         Width           =   1455
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Save"
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
      Left            =   5760
      TabIndex        =   1
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   960
      TabIndex        =   0
      Top             =   3960
      Width           =   4335
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   9360
      Top             =   6000
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
      Connect         =   $"BSIT_PAY.frx":0038
      OLEDBString     =   $"BSIT_PAY.frx":00D1
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select* from Table1"
      Caption         =   " "
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
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      DataField       =   "Sem"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   3000
      TabIndex        =   26
      Top             =   6360
      Width           =   1335
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      DataField       =   "RecAmount"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   4560
      TabIndex        =   25
      Top             =   6600
      Width           =   1095
   End
   Begin VB.Label Label11 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Type:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   5760
      TabIndex        =   24
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label Label12 
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
      Left            =   7200
      TabIndex        =   23
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      DataField       =   "Miscellanous Type"
      DataSource      =   "Adodc1"
      Height          =   1095
      Left            =   6240
      TabIndex        =   22
      Top             =   6000
      Width           =   1335
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      DataField       =   "TypeofPayment"
      DataSource      =   "Adodc1"
      Height          =   735
      Left            =   1920
      TabIndex        =   21
      Top             =   5880
      Width           =   2295
   End
   Begin VB.Label Label8 
      BackColor       =   &H00004000&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   720
      TabIndex        =   15
      Top             =   5040
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackColor       =   &H00004000&
      Caption         =   "Your Balance is:"
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
      Left            =   720
      TabIndex        =   14
      Top             =   4560
      Width           =   2535
   End
   Begin VB.Label Label3 
      DataField       =   "Balance"
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
      Left            =   3240
      TabIndex        =   13
      Top             =   4560
      Width           =   2055
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   5760
      TabIndex        =   12
      Top             =   240
      Width           =   2775
   End
   Begin VB.Label Label2 
      BackColor       =   &H00004000&
      Caption         =   "Current Balance:"
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
      Left            =   720
      TabIndex        =   11
      Top             =   360
      Width           =   2415
   End
   Begin VB.Label Label5 
      DataField       =   "Balance"
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
      Left            =   3360
      TabIndex        =   10
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
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
      Height          =   495
      Left            =   6120
      TabIndex        =   9
      Top             =   5040
      Visible         =   0   'False
      Width           =   2415
   End
End
Attribute VB_Name = "BSIT_PAY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
If Check1 = Checked Then
    Option3 = False
    Combo1.Visible = True
        BSIT_BALANCE.Label12.Visible = True
    Else
        Combo1.Visible = False
        BSIT_BALANCE.Label12.Visible = False
        BSIT_BALANCE.Label10.Caption = Option3.Caption
        End If


End Sub

Private Sub Command1_Click()
BSIT_PAY.Hide
Form1.Show

End Sub

Private Sub Command1_KeyPress(KeyAscii As Integer)
If Chr(KeyAscii) = Chr(27) Then
Call Command1_Click
End If

End Sub

Private Sub Command2_Click()
Dim Diff As Double
Dim Amount As Double
Adodc1.RecordSource = "select * from Table1 where Balance"
Adodc1.Refresh
             
Amount = Val(Text1.Text)
Balance = Adodc1.Recordset![Balance]
Diff = Balance - Amount
If Option1 = True Then
    Label4.Caption = "************************************" + vbNewLine + "" & Space(18) & "Paid" + vbCrLf + "************************************" & Space(15) & Option1.Caption + vbNewLine + vbNewLine + "Date:" + vbNewLine + FormatDateTime(Now, vbLongDate) & Space(5) & "Time:" + vbNewLine + FormatDateTime(Now, vbLongTime) + vbNewLine + vbNewLine + "Amount = " & Text1.Text + vbNewLine + "Your Balance = " & Balance & vbNewLine + "***********************************************" + vbNewLine + "Your New Balance: " & Diff

    ElseIf Option2 = True Then
    Label4.Caption = "************************************" + vbNewLine + "" & Space(18) & "Paid" + vbCrLf + "************************************" & Space(15) & Option2.Caption + vbNewLine + vbNewLine + "Date:" + vbNewLine + FormatDateTime(Now, vbLongDate) & Space(5) & "Time:" + vbNewLine + FormatDateTime(Now, vbLongTime) + vbNewLine + vbNewLine + "Amount = " & Text1.Text + vbNewLine + "Your Balance = " & Balance & vbNewLine + "***********************************************" + vbNewLine + "Your New Balance: " & Diff
    ElseIf Option1 = False Then

    Label3.Caption = ""
    MsgBox "Please Select Sem", vbExclamation

    End If

If Text1.Text = "" Then
    Label3.Caption = ""
    Label4.Caption = ""
    MsgBox " Please Enter The Amount for the payment", vbExclamation

        ElseIf Text1.Text <= 0 Then
            Label3.Caption = ""
            Label4.Caption = ""
            MsgBox "The amount you Entered is Unavailable", vbExclamation


                ElseIf Balance < Amount Then
                    MsgBox " Insufficient Balance", vbCritical
                    Label3.Caption = ""
                    Label4.Caption = ""

                        Else
                            Adodc1.Recordset.Update
                            Label3.Caption = Diff
                            Label7.Caption = FormatDateTime(Now, vbLongDate)
                            Label8.Caption = FormatDateTime(Now, vbLongTime)
If Option1 = True Then
    Label4.Caption = "************************************" + vbNewLine + "" & Space(18) & "Paid" + vbCrLf + "************************************" & Space(15) & Option1.Caption + vbNewLine + vbNewLine + "Date:" + vbNewLine + FormatDateTime(Now, vbLongDate) & Space(5) & "Time:" + vbNewLine + FormatDateTime(Now, vbLongTime) + vbNewLine + vbNewLine + "Amount = " & Text1.Text + vbNewLine + "Your Balance = " & Balance & vbNewLine + "***********************************************" + vbNewLine + "Your New Balance: " & Diff

        ElseIf Option2 = True Then
            Label4.Caption = "************************************" + vbNewLine + "" & Space(18) & "Paid" + vbCrLf + "************************************" & Space(15) & Option2.Caption + vbNewLine + vbNewLine + "Date:" + vbNewLine + FormatDateTime(Now, vbLongDate) & Space(5) & "Time:" + vbNewLine + FormatDateTime(Now, vbLongTime) + vbNewLine + vbNewLine + "Amount = " & Text1.Text + vbNewLine + "Your Balance = " & Balance & vbNewLine + "***********************************************" + vbNewLine + "Your New Balance: " & Diff
            ElseIf Option1 = False Then
            Label3.Caption = ""


                End If
End If

If Option3 = True Then
    Label9.Caption = Option3.Caption
    Label12.Caption = Option3.Caption
        ElseIf Check1 = Checked Then

            Label9.Caption = Check1.Caption
            Label12.Caption = Check1.Caption
            Label10.Caption = Combo1.Text
                 Else
                    Label4.Caption = ""
                    Label3.Caption = ""
                    MsgBox "Please Select Type Of payment", vbExclamation
   End If

    Label13.Caption = Text1.Text
      If Option1 = True Then
 Label14.Caption = Option1.Caption
 Else
  Label14.Caption = Option2.Caption
  End If
    
End Sub

Private Sub Command2_KeyPress(KeyAscii As Integer)
If Chr(KeyAscii) = Chr(13) Then
Call Command2_Click
End If
End Sub

Private Sub Command3_Click()
If Text1.Text = "" Then
MsgBox " Unable to Save This Field!", vbCritical
Label3.Caption = ""

ElseIf Label3.Caption = "" Then
MsgBox "Unable to save", vbCritical

Else
BSIT_BALANCE.Label1.Caption = Label3.Caption
BSIT_BALANCE.Label5.Caption = Label4.Caption
Adodc1.Recordset.AddNew
Label3.Caption = Diff
Label7.Caption = FormatDateTime(Now, vbLongDate)
Label8.Caption = FormatDateTime(Now, vbLongTime)
 Label13.Caption = Text1.Text
   If Option1 = True Then
 Label14.Caption = Option1.Caption
 Else
  Label14.Caption = Option2.Caption
  
  BSIT_PAY.Hide
BSIT_CODE.Show
BSIT_CODE.Text1.Text = ""



If Check1 = Checked Then
 Label9.Caption = Check1.Caption
 Label10.Caption = Combo1.Text
   Else
    Label9.Caption = Option3.Caption
    

If Check1 = Checked Then
 Label9.Caption = Check1.Caption
 Label10.Caption = Combo1.Text
 BSIT_BALANCE.Label13.Caption = Check1.Caption
 BSIT_BALANCE.Label15.Caption = Combo1.Text
   Else
   BSIT_BALANCE.Label13.Caption = Option1.Caption
    Label9.Caption = Option3.Caption
    BSIT_BALANCE.Label13.Caption = Option3.Caption
    End If
   
Adodc1.Recordset.AddNew
End If
End If
End If

End Sub

Private Sub Command4_Click()
Label4.Caption = ""
Label3.Caption = ""
Text1.Text = ""
Option1 = False
Option2 = False
Option3 = False
Option4 = False
Label12.Caption = ""
Check1 = 0

End Sub

Private Sub Form_Load()
Label3.Caption = ""
Label7.Caption = FormatDateTime(Now, vbLongDate)
Label8.Caption = FormatDateTime(Now, vbLongTime)
Adodc1.Recordset.Update
End Sub

Private Sub Option3_Click()
If Option3 = True Then
Check1 = 0
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
End Sub
