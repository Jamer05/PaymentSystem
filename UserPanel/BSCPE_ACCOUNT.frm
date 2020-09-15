VERSION 5.00
Begin VB.Form BSCPE_ACCOUNT 
   BackColor       =   &H00004000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BSCPE_Account"
   ClientHeight    =   5085
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7980
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   7980
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H0000C000&
      BorderStyle     =   0  'None
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
      Height          =   2175
      Left            =   720
      TabIndex        =   0
      Top             =   2520
      Width           =   6375
      Begin VB.CommandButton Command4 
         Caption         =   "Logout"
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
         Left            =   2520
         TabIndex        =   4
         Top             =   1560
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Pay"
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
         Left            =   360
         TabIndex        =   3
         Top             =   840
         Width           =   1815
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Check Balance"
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
         Left            =   2160
         TabIndex        =   2
         Top             =   840
         Width           =   2295
      End
      Begin VB.CommandButton Command3 
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
         Height          =   495
         Left            =   4440
         TabIndex        =   1
         Top             =   840
         Width           =   1815
      End
   End
   Begin VB.Label BSIT_ACCOUNT 
      BackStyle       =   0  'Transparent
      Caption         =   "MENU"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   48.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   2520
      TabIndex        =   5
      Top             =   720
      Width           =   3135
   End
End
Attribute VB_Name = "BSCPE_ACCOUNT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
BSCPE_ACCOUNT.Hide
BSCPE_PAY.Show
BSCPE_PAY.Label3.Caption = ""
BSCPE_PAY.Text1.Text = ""
BSCPE_PAY.Label4.Caption = ""
BSCPE_PAY.Option1 = False
BSCPE_PAY.Option2 = False
BSCPE_PAY.Option3 = False
BSCPE_PAY.Check1 = False
BSCPE_PAY.Text1.SetFocus

End Sub

Private Sub Command2_Click()
BSCPE_ACCOUNT.Hide
BSCPE_BALANCE.Show

If BSCPE_PAY.Label4.Caption = "" Then
BSCPE_BALANCE.Label5.Caption = vbNewLine + vbNewLine + vbNewLine + Space(10) & "Nothing" & Space(5) & "No History"
End If
If BSCPE_BALANCE.Label10.Caption = "Miscellanous" Then
BSCPE_BALANCE.Label12.Visible = True
Else
 BSCPE_BALANCE.Label12.Visible = False
 End If
  If BSCPE_PAY.Check1 = 1 Then
 BSCPE_PAY.Label9.Caption = BSCPE_PAY.Check1.Caption
 BSCPE_PAY.Label10.Caption = BSCPE_PAY.Combo1.Text
 BSCPE_BALANCE.Label13.Caption = BSCPE_PAY.Check1.Caption
 BSCPE_BALANCE.Label15.Caption = BSCPE_PAY.Combo1.Text
   ElseIf BSIT_PAY.Option3 = True Then
   BSCPE_BALANCE.Label13.Caption = BSCPE_PAY.Option1.Caption
     BSCPE_PAY.Label9.Caption = BSCPE_PAY.Option3.Caption
    BSCPE_BALANCE.Label13.Caption = BSCPE_PAY.Option3.Caption
    End If
If BSCPE_BALANCE.Label13.Caption = "Miscellanous" Then
BSCPE_BALANCE.Label15.Visible = True
BSCPE_BALANCE.Label16.Visible = True
End If
End Sub

Private Sub Command3_Click()
BSCPE_ACCOUNT.Hide
Course.Show
Course.Combo1.Text = ""
End Sub

Private Sub Command4_Click()
BSCPE_ACCOUNT.Hide
BSCPE_LOGIN.Show
BSCPE_LOGIN.Text1.Text = ""
BSCPE_LOGIN.Text2.Text = ""
BSCPE_LOGIN.Text2.SetFocus
End Sub
