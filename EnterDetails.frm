VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form EnterDetails 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Your software title goes here."
   ClientHeight    =   3510
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
   ScaleHeight     =   3510
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cd 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "TUX license file (*.tux)|*.tux"
   End
   Begin VB.CommandButton cmdLtUX 
      Caption         =   "..."
      Height          =   255
      Left            =   6360
      TabIndex        =   10
      Top             =   2400
      Width           =   735
   End
   Begin VB.TextBox txtTuxL 
      Height          =   285
      Left            =   120
      TabIndex        =   9
      Top             =   2400
      Width           =   6135
   End
   Begin VB.TextBox txtFingerprint 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   360
      Width           =   6975
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "&Cancel"
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
      Left            =   4680
      TabIndex        =   5
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton CmdReg 
      Caption         =   "&Register"
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
      Left            =   1200
      TabIndex        =   4
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox txtUsername 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   6975
   End
   Begin VB.TextBox txtGen 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   6975
   End
   Begin VB.Label lblTUXL 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&License file (TUX):"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2160
      Width           =   6975
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   7080
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Label lblFprint 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&Hardware Fingerprint:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   6975
   End
   Begin VB.Label Label1 
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
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Registration Code:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   1815
   End
End
Attribute VB_Name = "EnterDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim clsDS2 As New clsDS2

Private Sub cmdLtUX_Click()

    cd.ShowOpen
    
    txtTuxL.Text = cd.FileName
    

End Sub

'*********************************************'
'                                             '
' SimpleTrial                                 '
' Feel free to re-distrubute this code, since '
' this code is freeware :).                   '
'                                             '
' Please vote for me.                         '
'                                             '
'*********************************************'

Private Sub cmdQuit_Click()

    'Quit the form when a user decides to.
        Me.Hide

End Sub

Private Sub CmdReg_Click()
    
    If KeyGen((txtFingerprint & txtUsername), "%%%%-0044-0X7Q-PKQ88-BQW91-%%%%", 2) = txtGen Then CheckLicense Else MsgBox "Registration failed!", vbCritical, "Failed!"

End Sub

Public Sub CheckLicense()

    Dim CheckValid As String
    Dim H_FINGERPRINT As String
    Dim USERName As String
    Dim SERIAL As String
    
    On Error GoTo ErrHandle:
    
    If Len(txtUsername) < 1 Then Exit Sub

    Open (txtTuxL) For Input As #1
        Line Input #1, CheckValid
        Line Input #1, H_FINGERPRINT
        Line Input #1, USERName
    Close #1
    
    CheckValid = EnigmaDecrypt(CheckValid)
    H_FINGERPRINT = EnigmaDecrypt(H_FINGERPRINT)
    USERName = EnigmaDecrypt(USERName)
    If CheckValid = "[TRIAL CONFIG FILE]" Then GoTo regPt2
    
ErrHandle:
    MsgBox "Registration failed!", vbCritical, "Failed!": End
regPt2:
    If H_FINGERPRINT = txtFingerprint Then GoTo regPt3
regPt3:
    If USERName = txtUsername Then GoTo regPt4
regPt4:
    If KeyGen((H_FINGERPRINT & USERName), "%%%%-0044-0X7Q-PKQ88-BQW91-%%%%", 2) = SERIAL Then GoTo regpt5
regpt5:
    'Encrypt the file to stop people from looking at this hidden info.
        txtFingerprint.Text = clsDS2.EncryptString(txtFingerprint, "%%%%-0044-0X7Q-PKQ88-BQW91-%%%%", True)
        txtUsername.Text = clsDS2.EncryptString(txtUsername.Text, "%%%%-0044-0X7Q-PKQ88-BQW91-%%%%", True)
        txtGen.Text = clsDS2.EncryptString(txtGen.Text, "%%%%-0044-0X7Q-PKQ88-BQW91-%%%%", True)
    
    'Write the details to file, if they are correct then the software will be registered.
    
    Open "C:\WINDOWS\system32\hlgxu.002" For Output As #2
        Print #2, EnigmaEncrypt(txtFingerprint)
        Print #2, EnigmaEncrypt(txtUsername)
        Print #2, EnigmaEncrypt(txtGen)
    Close #2
    
    Me.Hide
    
    'Copy the temp file to the trial config file.
    FileCopy "C:\WINDOWS\system32\hlgxu.002", "C:\WINDOWS\system32\hlgxu.001"
    Kill "C:\WINDOWS\system32\hlgxu.002"
    
    MsgBox "Registration Successfull!" & vbCrLf & vbCrLf & "Application needs to be restarted for changes to take effect!..", vbInformation: Unload Me
    
End Sub

Private Sub Form_Load()

    txtFingerprint = KeyGen(GetSystemSerial, "^£_[UNKNOWN]_&*", 3)

End Sub
