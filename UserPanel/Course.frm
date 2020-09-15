VERSION 5.00
Begin VB.Form Course 
   BackColor       =   &H00004000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   8385
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11355
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8385
   ScaleWidth      =   11355
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   8415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11415
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   3015
         Left            =   840
         TabIndex        =   8
         Top             =   3720
         Width           =   9735
         Begin VB.Image Image3 
            Height          =   2775
            Left            =   3480
            Picture         =   "Course.frx":0000
            Stretch         =   -1  'True
            Top             =   120
            Width           =   3015
         End
         Begin VB.Image Image1 
            Height          =   2775
            Left            =   6720
            Picture         =   "Course.frx":2896
            Stretch         =   -1  'True
            Top             =   120
            Width           =   3015
         End
         Begin VB.Image Image2 
            Height          =   2715
            Left            =   120
            Picture         =   "Course.frx":8A6E
            Stretch         =   -1  'True
            Top             =   120
            Width           =   3075
         End
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H0000FF00&
         Cancel          =   -1  'True
         Caption         =   "BSBA"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   8040
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   6720
         Width           =   2415
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H0000FF00&
         Caption         =   "BSCPE"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   6720
         Width           =   2535
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H0000FF00&
         Caption         =   "BSIT"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   6720
         UseMaskColor    =   -1  'True
         Width           =   2415
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Continue"
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
         Left            =   7440
         TabIndex        =   4
         Top             =   3120
         Width           =   1335
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Course.frx":DA72
         Left            =   4680
         List            =   "Course.frx":DA7F
         TabIndex        =   1
         Top             =   3120
         Width           =   2295
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "IETI COLLEGE ALABANG Inc..."
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   42
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   1095
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   11415
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Select/Type"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   16.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   495
         Left            =   4680
         TabIndex        =   9
         Top             =   2520
         Width           =   2295
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000C000&
         Caption         =   "Select"
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
         Left            =   3240
         TabIndex        =   3
         Top             =   3120
         Width           =   975
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000C000&
         Caption         =   "Choose Course"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   4080
         TabIndex        =   2
         Top             =   1920
         Width           =   3615
      End
   End
End
Attribute VB_Name = "Course"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command3_Click()
Course.Hide
BSCPE_LOGIN.Show
BSCPE_LOGIN.Text1.Text = ""
BSCPE_LOGIN.Text2.Text = ""

BSCPE_LOGIN.Text2.SetFocus
End Sub

Private Sub Command4_Click()
Course.Hide
BSBA_LOGIN.Show
BSBA_LOGIN.Text1.Text = ""
BSBA_LOGIN.Text2.Text = ""
BSBA_LOGIN.Text1.SetFocus
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)

If Chr(KeyAscii) = Chr(13) Then
Call Command1_Click

End If
End Sub

Private Sub Command1_Click()
Combo1.SetFocus
If Combo1.Text = "BSIT" Then
Course.Hide
BSIT_LOGIN.Show
BSIT_LOGIN.Text1.SetFocus
BSIT_LOGIN.Text1.Text = ""
BSIT_LOGIN.Text1.Text = ""

ElseIf Combo1.Text = "bsit" Then
Course.Hide
BSIT_LOGIN.Show
BSIT_LOGIN.Text1.SetFocus
BSIT_LOGIN.Text1.Text = ""
BSIT_LOGIN.Text1.Text = ""
ElseIf Combo1.Text = "BSBA" Then
Course.Hide
BSBA_LOGIN.Show
BSBA_LOGIN.Text1.SetFocus
BSBA_LOGIN.Text1.Text = ""
BSBA_LOGIN.Text2.Text = ""
ElseIf Combo1.Text = "bsba" Then
Course.Hide
BSBA_LOGIN.Show
BSBA_LOGIN.Text1.SetFocus
BSBA_LOGIN.Text1.Text = ""
BSBA_LOGIN.Text2.Text = ""
ElseIf Combo1.Text = "BSCPE" Then
Course.Hide
BSCPE_LOGIN.Show
BSCPE_LOGIN.Text2.SetFocus
BSCPE_LOGIN.Text1.Text = ""
BSCPE_LOGIN.Text2.Text = ""
ElseIf Combo1.Text = "bscpe" Then
Course.Hide
BSCPE_LOGIN.Show
BSCPE_LOGIN.Text2.SetFocus
BSCPE_LOGIN.Text1.Text = ""
BSCPE_LOGIN.Text2.Text = ""
Else
Combo1.Text = ""
MsgBox "Please enter the only available course", vbExclamation
End If
BSIT_LOGIN.Text1 = ""
BSIT_LOGIN.Text2 = ""
BSBA_LOGIN.Text1 = ""
BSBA_LOGIN.Text2 = ""
BSCPE_LOGIN.Text1 = ""
BSCPE_LOGIN.Text2 = ""
End Sub

Private Sub Command2_Click()
Course.Hide
BSIT_LOGIN.Show
BSIT_LOGIN.Text1.Text = ""
BSIT_LOGIN.Text2.Text = ""
BSIT_LOGIN.Text1.SetFocus

End Sub

