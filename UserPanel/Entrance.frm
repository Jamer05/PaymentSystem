VERSION 5.00
Begin VB.Form Entrance 
   BackColor       =   &H00004000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   7545
   ClientLeft      =   150
   ClientTop       =   780
   ClientWidth     =   9300
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   9300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00008000&
      Height          =   6855
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   8895
      Begin VB.CommandButton Command1 
         Caption         =   "Continue"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   7200
         TabIndex        =   6
         Top             =   6240
         Width           =   1455
      End
      Begin VB.Label Label5 
         BackColor       =   &H0000C000&
         Caption         =   "BSCPE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   6720
         TabIndex        =   5
         Top             =   5760
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackColor       =   &H0000C000&
         Caption         =   "BSIT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   4080
         TabIndex        =   4
         Top             =   5760
         Width           =   855
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000C000&
         Caption         =   "BSBA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1440
         TabIndex        =   3
         Top             =   5760
         Width           =   975
      End
      Begin VB.Image Image3 
         Height          =   2400
         Left            =   5880
         Picture         =   "Entrance.frx":0000
         Stretch         =   -1  'True
         Top             =   3360
         Width           =   2640
      End
      Begin VB.Image Image2 
         Height          =   2355
         Left            =   3120
         Picture         =   "Entrance.frx":2896
         Stretch         =   -1  'True
         Top             =   3360
         Width           =   2475
      End
      Begin VB.Image Image1 
         Height          =   2355
         Left            =   480
         Picture         =   "Entrance.frx":789A
         Stretch         =   -1  'True
         Top             =   3360
         Width           =   2475
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000C000&
         Caption         =   "    Computerized Payment System"
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
         Height          =   735
         Left            =   240
         TabIndex        =   2
         Top             =   1200
         Width           =   8415
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000C000&
         Caption         =   "              Course Available"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   29.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   3615
         Left            =   240
         TabIndex        =   1
         Top             =   2760
         Width           =   8415
      End
   End
   Begin VB.Menu About 
      Caption         =   "About_Us"
   End
End
Attribute VB_Name = "Entrance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Private Sub About_Click()
Entrance.Hide
AboutUs.Show
End Sub

Private Sub Command1_Click()
Entrance.Hide
Course.Show

End Sub

Private Sub Command1_KeyPress(KeyAscii As Integer)
If Chr(KeyAscii) = Chr(13) Then
Call Command1_Click
End If
End Sub

