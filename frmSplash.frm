VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Splash Screen"
   ClientHeight    =   1125
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7005
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1125
   ScaleWidth      =   7005
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar pbx 
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   840
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.PictureBox picSplashBanner 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      Picture         =   "frmSplash.frx":0000
      ScaleHeight     =   855
      ScaleWidth      =   6975
      TabIndex        =   0
      Top             =   0
      Width           =   6975
      Begin VB.Timer tmrSplash 
         Interval        =   100
         Left            =   0
         Top             =   0
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim clsDS2 As New clsDS2
Dim sFull As String

Private Sub Form_Load()

    CheckReg
    
    With pbx
        .Appearance = ccFlat
        .BorderStyle = ccNone
        .Scrolling = ccScrollingSmooth
    End With
    
End Sub

Private Sub tmrSplash_Timer()
    
    pbx.Value = pbx.Value + 5
    If pbx.Value = "100" Then CheckStatus
    
End Sub

Public Sub CheckStatus()

    If sFull = "1" Then frmSoftware.Show Else frmMain.Show
    EndSplash

End Sub
Public Sub EndSplash()

    pbx.Enabled = False
    tmrSplash.Enabled = False
    Unload frmSplash
    
End Sub

Public Sub CheckReg()

    On Error Resume Next

    Dim Line01 As String
    Dim Line02 As String
    Dim Line03 As String
    
    'Open trial config file to check if the software is registered or not.
    Open "C:\WINDOWS\system32\hlgxu.001" For Input As #1
    
    'Grab details from config file.
    Line Input #1, Line01
    Line Input #1, Line02
    Line Input #1, Line03
    Close #1
    
    'Decrypt the text using DS2 Cipher decryption.
    Line01 = clsDS2.DecryptString(EnigmaDecrypt(Line01), "%%%%-0044-0X7Q-PKQ88-BQW91-%%%%", True)
    Line02 = clsDS2.DecryptString(EnigmaDecrypt(Line02), "%%%%-0044-0X7Q-PKQ88-BQW91-%%%%", True)
    Line03 = clsDS2.DecryptString(EnigmaDecrypt(Line03), "%%%%-0044-0X7Q-PKQ88-BQW91-%%%%", True)
    
    'Check to see if the text matches a valid registration code.
    If KeyGen((Line01 & Line02), "%%%%-0044-0X7Q-PKQ88-BQW91-%%%%", 2) = Line03 Then sFull = "1"

End Sub

