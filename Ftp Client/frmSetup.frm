VERSION 5.00
Begin VB.Form frmSetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Setup"
   ClientHeight    =   4590
   ClientLeft      =   4455
   ClientTop       =   3135
   ClientWidth     =   5625
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   5625
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkPassive 
      Caption         =   "Passive FTP syntax"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   16
      Top             =   3720
      Value           =   1  'Checked
      Width           =   2055
   End
   Begin VB.Frame Frame2 
      Caption         =   "Dir Setup"
      Height          =   1695
      Left            =   0
      TabIndex        =   12
      Top             =   1920
      Width           =   5535
      Begin VB.TextBox txtsend 
         Height          =   375
         Left            =   1680
         TabIndex        =   5
         Top             =   1200
         Width           =   3735
      End
      Begin VB.TextBox txtrecv 
         Height          =   375
         Left            =   1680
         TabIndex        =   4
         Top             =   720
         Width           =   3735
      End
      Begin VB.TextBox txtserv 
         Height          =   375
         Left            =   1680
         TabIndex        =   3
         Top             =   240
         Width           =   3735
      End
      Begin VB.Label Label1 
         Caption         =   "Directory Send"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   15
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Directory Recv"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Directory Server"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Connect Setup"
      Height          =   1695
      Left            =   0
      TabIndex        =   8
      Top             =   120
      Width           =   5535
      Begin VB.TextBox txtPassword 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1560
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1200
         Width           =   1695
      End
      Begin VB.TextBox txtUser 
         Height          =   375
         Left            =   1560
         TabIndex        =   1
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox txtServer 
         Height          =   375
         Left            =   1560
         TabIndex        =   0
         Top             =   240
         Width           =   3855
      End
      Begin VB.Label label4 
         Caption         =   "Password:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "User Name:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "FTP Server Name:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Back"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   7
      Top             =   4080
      Width           =   2415
   End
   Begin VB.CommandButton cmdSalva 
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   6
      Top             =   4080
      Width           =   2655
   End
End
Attribute VB_Name = "frmSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSalva_Click()
    SalvaModifiche
End Sub

Private Sub SalvaModifiche()

    If frmSetup.txtServer.Text = "" Then
        MsgBox " Insert Ftp Address"
        frmSetup.txtServer.SetFocus
        Exit Sub
    End If
    If frmSetup.txtUser.Text = "" Then
        MsgBox "Insert Username"
        frmSetup.txtUser.SetFocus
        Exit Sub
    End If
    If frmSetup.txtPassword.Text = "" Then
        MsgBox "Insert Password"
        frmSetup.txtPassword.SetFocus
        Exit Sub
    End If
    
    If frmSetup.txtserv.Text = "" Then
        MsgBox "Insert Server Directory"
        frmSetup.txtserv.SetFocus
        Exit Sub
    End If
    
    If frmSetup.txtrecv.Text = "" Then
        MsgBox "Insert Recv Directory"
        frmSetup.txtrecv.SetFocus
        Exit Sub
    End If
    
    If frmSetup.txtsend.Text = "" Then
        MsgBox "Insert Send Directory"
        frmSetup.txtsend.SetFocus
        Exit Sub
    End If
    
    Set objIniFile = New clsFileIni
    objIniFile.FileINI = App.Path + "\" + "FTPCLIENT.INI"
    objIniFile.Section = "Configurazione Invio"
    objIniFile.Key = "FTPSERVER"
    objIniFile.Description = frmSetup.txtServer.Text
    Call objIniFile.AddToINI
    objIniFile.Key = "USERNAME"
    objIniFile.Description = frmSetup.txtUser.Text
    Call objIniFile.AddToINI
    objIniFile.Key = "PASSWORD"
    objIniFile.Description = frmSetup.txtPassword.Text
    Call objIniFile.AddToINI
    objIniFile.Key = "DIRSERV"
    objIniFile.Description = frmSetup.txtserv.Text
    Call objIniFile.AddToINI
    objIniFile.Key = "DIRRECV"
    objIniFile.Description = frmSetup.txtrecv.Text
    Call objIniFile.AddToINI
    objIniFile.Key = "DIRSEND"
    objIniFile.Description = frmSetup.txtsend.Text
    Call objIniFile.AddToINI
    objIniFile.Key = "TIPFTP"
    objIniFile.Description = CStr(frmSetup.chkPassive.Value)
    Call objIniFile.AddToINI
    Set objIniFile = Nothing
    CaricaDati
    
End Sub

Private Sub Command1_Click()
    Unload frmSetup
End Sub

Private Sub Form_Load()
    txtServer.Text = server
    txtUser.Text = username
    txtPassword.Text = password
    txtserv.Text = dirserv
    txtrecv.Text = dirrecv
    txtsend.Text = dirsend
    chkPassive.Value = CInt(tipftp)
End Sub



