VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00004000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BSIT_Account"
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8085
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   8085
   ShowInTaskbar   =   0   'False
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
      Left            =   840
      TabIndex        =   1
      Top             =   2280
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
         TabIndex        =   5
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
         TabIndex        =   4
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
         TabIndex        =   3
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
         TabIndex        =   2
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
      Left            =   2760
      TabIndex        =   0
      Top             =   600
      Width           =   3135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
BSIT_PAY.Label12.Caption = ""
BSIT_PAY.Show
Form1.Hide
BSIT_PAY.Label3.Caption = ""
BSIT_PAY.Text1.Text = ""
BSIT_PAY.Label4.Caption = ""
BSIT_PAY.Option1 = False
BSIT_PAY.Option2 = False
BSIT_PAY.Option3 = False
BSIT_PAY.Check1 = 0
BSIT_PAY.Text1.SetFocus

End Sub

Private Sub Command2_Click()
Form1.Hide
BSIT_BALANCE.Show
BSIT_BALANCE.Label13.Caption = ""
If BSIT_PAY.Label4.Caption = "" Then
BSIT_BALANCE.Label5.Caption = vbNewLine + vbNewLine + vbNewLine + Space(10) & "Nothing" & Space(5) & "No History"
End If
If BSIT_BALANCE.Label10.Caption = "Miscellanous" Then
BSIT_BALANCE.Label12.Visible = True
Else
BSIT_BALANCE.Label12.Visible = False
End If
 If BSIT_PAY.Check1 = Checked Then
 BSIT_PAY.Label9.Caption = BSIT_PAY.Check1.Caption
 BSIT_PAY.Label10.Caption = BSIT_PAY.Combo1.Text
 BSIT_BALANCE.Label13.Caption = BSIT_PAY.Check1.Caption
 BSIT_BALANCE.Label15.Caption = BSIT_PAY.Combo1.Text
   ElseIf BSIT_PAY.Option3 = True Then
   BSIT_BALANCE.Label13.Caption = BSIT_PAY.Option1.Caption
     BSIT_PAY.Label9.Caption = BSIT_PAY.Option3.Caption
    BSIT_BALANCE.Label13.Caption = BSIT_PAY.Option3.Caption
    End If
If BSIT_BALANCE.Label10.Caption = "Miscellanous" Then
BSIT_BALANCE.Label15.Visible = True
BSIT_BALANCE.Label16.Visible = True
End If
 

End Sub

Private Sub Command3_Click()
Course.Show
Form1.Hide
Course.Combo1.Text = ""

End Sub

Private Sub Command3_KeyPress(KeyAscii As Integer)
If Chr(KeyAscii) = Chr(27) Then
Call Command3_Click
End If
End Sub

Private Sub Command4_Click()
Form1.Hide
BSIT_LOGIN.Show
BSIT_LOGIN.Text1.Text = ""
BSIT_LOGIN.Text2.Text = ""
BSIT_LOGIN.Text1.SetFocus
End Sub

