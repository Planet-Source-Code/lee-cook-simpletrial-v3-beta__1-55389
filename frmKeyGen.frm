VERSION 5.00
Begin VB.Form frmKeyGen 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Software title goes here."
   ClientHeight    =   2760
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7215
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdMkLic 
      Appearance      =   0  'Flat
      Caption         =   "&Create license file.."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   7
      Top             =   2040
      Width           =   2175
   End
   Begin VB.TextBox txtFingerprint 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   6975
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   5
      Top             =   2040
      Width           =   2175
   End
   Begin VB.TextBox txtGen 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   6975
   End
   Begin VB.TextBox txtUsername 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   6975
   End
   Begin VB.Label lblFprint 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&Hardware Fingerprint:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   6975
   End
   Begin VB.Label lblRCode 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Registration Code:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   6975
   End
   Begin VB.Label lblUsr 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&Username:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   6975
   End
End
Attribute VB_Name = "frmKeyGen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'*********************************************'
'                                             '
' SimpleTrial                                 '
' Feel free to re-distrubute this code, since '
' this code is freeware :).                   '
'                                             '
' Please vote for me.                         '
'                                             '
'*********************************************'

Private Sub cmdExit_Click()

    'Quit the form.
         Me.Hide

End Sub

Private Sub cmdMkLic_Click()

    frmMkLic.Show: Me.Hide

End Sub

Private Sub Form_Load()

    txtFingerprint = KeyGen(GetSystemSerial, "^£_[UNKNOWN]_&*", 3)

End Sub

Private Sub txtUsername_Change()

    'Generate registration info.
        txtGen.Text = KeyGen((txtFingerprint & txtUsername), "%%%%-0044-0X7Q-PKQ88-BQW91-%%%%", 2)

End Sub
