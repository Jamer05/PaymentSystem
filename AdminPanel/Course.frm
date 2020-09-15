VERSION 5.00
Begin VB.Form Course 
   BackColor       =   &H00808000&
   Caption         =   "Course"
   ClientHeight    =   5205
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9270
   LinkTopic       =   "Form1"
   ScaleHeight     =   5205
   ScaleWidth      =   9270
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4575
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   8535
      Begin VB.CommandButton Command4 
         Caption         =   "Back"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   6240
         TabIndex        =   5
         Top             =   3360
         Width           =   1455
      End
      Begin VB.CommandButton Command3 
         Caption         =   "BSBA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   4560
         TabIndex        =   3
         Top             =   3360
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         Caption         =   "BSCPE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2880
         TabIndex        =   2
         Top             =   3360
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         Caption         =   "BSIT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1080
         TabIndex        =   1
         Top             =   3360
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Student Registration    /Choose Course"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   39
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   2775
         Left            =   720
         TabIndex        =   4
         Top             =   360
         Width           =   7575
      End
   End
End
Attribute VB_Name = "Course"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Course.Hide
BSIT.Show
End Sub

Private Sub Command2_Click()
Menu.Hide
BSCPE.Show
End Sub

Private Sub Command3_Click()
Menu.Hide
BSBA.Show

End Sub

Private Sub Command4_Click()
Course.Hide
Menu.Show
End Sub
