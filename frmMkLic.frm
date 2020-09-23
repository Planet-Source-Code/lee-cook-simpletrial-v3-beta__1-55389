VERSION 5.00
Begin VB.Form frmMkLic 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "License file generator."
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7215
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   4440
      TabIndex        =   7
      Top             =   2520
      Width           =   2055
   End
   Begin VB.CommandButton cmdMkLic 
      Caption         =   "&Create license file.."
      Height          =   495
      Left            =   720
      TabIndex        =   6
      Top             =   2520
      Width           =   2055
   End
   Begin VB.TextBox txtRCode 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Top             =   2040
      Width           =   6975
   End
   Begin VB.TextBox txtUsername 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   6975
   End
   Begin VB.TextBox txtFingerprint 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   6975
   End
   Begin VB.Label lblRCode 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&Registration Code:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   6975
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   7080
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Label lblUsr 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&Username:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   6975
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   7080
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label lblFprint 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&Hardware Fingerprint:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6975
   End
End
Attribute VB_Name = "frmMkLic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()

    Me.Hide

End Sub

Private Sub cmdMkLic_Click()

    Open App.Path & "\license.tux" For Output As #1
        Print #1, EnigmaEncrypt("[TRIAL CONFIG FILE]")
        Print #1, EnigmaEncrypt(txtFingerprint)
        Print #1, EnigmaEncrypt(txtUsername)
        Print #1, EnigmaEncrypt(txtRCode)
        Print #1, EnigmaEncrypt(Time)
        Print #1, EnigmaEncrypt(Date)
        Print #1, EnigmaEncrypt("TUX Encrypter")
        MsgBox "License file created successfully..", vbInformation, "Creation wizard."
    Close #1
End Sub

Private Sub txtFingerprint_Change()

    txtRCode = KeyGen((txtFingerprint & txtUsername), "%%%%-0044-0X7Q-PKQ88-BQW91-%%%%", 2)

End Sub

Private Sub txtUsername_Change()

    txtRCode = KeyGen((txtFingerprint & txtUsername), "%%%%-0044-0X7Q-PKQ88-BQW91-%%%%", 2)

End Sub
