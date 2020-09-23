VERSION 5.00
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Software title goes here."
   ClientHeight    =   1635
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5190
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1635
   ScaleWidth      =   5190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TmrMain 
      Interval        =   10000
      Left            =   240
      Top             =   840
   End
   Begin VB.CommandButton CmdAbout 
      Caption         =   "About"
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
      Left            =   2640
      TabIndex        =   4
      Top             =   840
      Width           =   1575
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
      Left            =   960
      TabIndex        =   3
      Top             =   840
      Width           =   1575
   End
   Begin VB.CommandButton CmdEnter 
      Caption         =   "Enter Trial"
      Enabled         =   0   'False
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
      Left            =   1800
      TabIndex        =   1
      Top             =   240
      Width           =   1575
   End
   Begin VB.CommandButton cmdKGen 
      Caption         =   "Key Generator"
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
      Left            =   3480
      TabIndex        =   2
      Top             =   240
      Width           =   1575
   End
   Begin VB.CommandButton CmdEntSerial 
      Caption         =   "Enter Serial"
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
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "frmMain"
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

Dim clsDS2 As New clsDS2

Private Sub CmdAbout_Click()

    'Show details about your software.
        MsgBox "Company Name: " & App.CompanyName & vbCrLf & "Product Name: " & App.ProductName & vbCrLf & "Version: " & App.Major & "." & App.Revision & "." & App.Minor & vbCrLf & vbCrLf & "Little message about your product here.."

End Sub

Private Sub cmdEnter_Click()

    'Load the software.
        frmSoftware.Show: Unload Me
        
    'Add the unregistered status to the software.
        frmSoftware.Caption = "" & App.ProductName & " (Unregistered Version)"

End Sub

Private Sub cmdExit_Click()

    'Terminate the program if the user decides to.
        End

End Sub

Private Sub CmdKGen_Click()

    'Load the Key Generator form.
        frmKeyGen.Show

End Sub

Private Sub CmdEntSerial_Click()

    'Load the details entry form.
        EnterDetails.Show
End Sub

Private Sub TmrMain_Timer()
    
    'Delay the enter button.
        CmdEnter.Enabled = True

End Sub
