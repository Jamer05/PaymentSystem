VERSION 5.00
Begin VB.Form BSBA_ACCOUNT 
   BackColor       =   &H00004000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BSBA_Account"
   ClientHeight    =   5085
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7905
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   7905
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
      Top             =   2160
      Width           =   6375
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
         TabIndex        =   4
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
         TabIndex        =   3
         Top             =   840
         Width           =   2295
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
         TabIndex        =   2
         Top             =   840
         Width           =   1815
      End
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
         TabIndex        =   1
         Top             =   1560
         Width           =   1695
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
      Top             =   480
      Width           =   3135
   End
End
Attribute VB_Name = "BSBA_ACCOUNT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
BSBA_PAY.Label4.Caption = ""
BSBA_ACCOUNT.Hide
BSBA_PAY.Show
BSBA_PAY.Label3.Caption = ""
BSBA_PAY.Text1.Text = ""
BSBA_PAY.Label4.Caption = ""
BSBA_PAY.Option1 = False
BSBA_PAY.Option2 = False
BSBA_PAY.Option3 = False
BSBA_PAY.Check1 = False
BSBA_PAY.Text1.SetFocus
End Sub

Private Sub Command2_Click()
BSBA_ACCOUNT.Hide
BSBA_BALANCE.Show

If BSBA_PAY.Label4.Caption = "" Then
BSBA_BALANCE.Label5.Caption = vbNewLine + vbNewLine + vbNewLine + Space(10) & "Nothing" & Space(5) & "No History"
End If
If BSBA_BALANCE.Label10.Caption = "Miscellanous" Then
BSBA_BALANCE.Label12.Visible = True
Else
 BSBA_BALANCE.Label12.Visible = False
 End If
  If BSBA_PAY.Check1 = Checked Then
 BSBA_PAY.Label9.Caption = BSBA_PAY.Check1.Caption
 BSBA_PAY.Label10.Caption = BSBA_PAY.Combo1.Text
 BSBA_BALANCE.Label13.Caption = BSBA_PAY.Check1.Caption
 BSBA_BALANCE.Label15.Caption = BSBA_PAY.Combo1.Text
   ElseIf BSIT_PAY.Option3 = True Then
   BSBA_BALANCE.Label13.Caption = BSBA_PAY.Option1.Caption
     BSBA_PAY.Label9.Caption = BSBA_PAY.Option3.Caption
    BSBA_BALANCE.Label13.Caption = BSBA_PAY.Option3.Caption
    End If
If BSBA_BALANCE.Label13.Caption = "Miscellanous" Then
BSBA_BALANCE.Label15.Visible = True
BSBA_BALANCE.Label15.Visible = True
End If

End Sub

Private Sub Command3_Click()
BSBA_ACCOUNT.Hide
Course.Show
Course.Combo1.Text = ""

End Sub

Private Sub Command4_Click()
BSBA_ACCOUNT.Hide
BSBA_LOGIN.Show
BSBA_LOGIN.Text1.Text = ""
BSBA_LOGIN.Text2.Text = ""
BSBA_LOGIN.Text1.SetFocus
End Sub

