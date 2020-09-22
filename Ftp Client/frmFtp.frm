VERSION 5.00
Begin VB.Form frmFTP 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "FTP Client"
   ClientHeight    =   5610
   ClientLeft      =   3060
   ClientTop       =   2505
   ClientWidth     =   8475
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   8475
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   1215
      Index           =   1
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   8415
      Begin VB.CommandButton cmdPulsante 
         Caption         =   "&Connect"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   0
         Left            =   120
         Picture         =   "frmFtp.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   240
         Width           =   2055
      End
      Begin VB.CommandButton cmdPulsante 
         Caption         =   "&Disconnect"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   1
         Left            =   2280
         Picture         =   "frmFtp.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmdPulsante 
         Caption         =   "&Setup"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   2
         Left            =   4080
         Picture         =   "frmFtp.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   240
         Width           =   2055
      End
      Begin VB.CommandButton cmdPulsante 
         Caption         =   "&Automatic"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   3
         Left            =   6240
         Picture         =   "frmFtp.frx":0CC6
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Pulsante di Uscita"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   2
      Left            =   0
      TabIndex        =   6
      Top             =   4320
      Width           =   2895
      Begin VB.CommandButton cmdPulsante 
         Caption         =   "&Exit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   4
         Left            =   120
         Picture         =   "frmFtp.frx":0FD0
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Pulsanti di Avvio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   3
      Left            =   3000
      TabIndex        =   5
      Top             =   4320
      Width           =   5415
      Begin VB.CommandButton cmdPulsante 
         Caption         =   "&UpLoad"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   6
         Left            =   2280
         Picture         =   "frmFtp.frx":12DA
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdPulsante 
         Caption         =   "&DownLoad"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   7
         Left            =   3840
         Picture         =   "frmFtp.frx":15E4
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdPulsante 
         Caption         =   "&Log"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   5
         Left            =   120
         Picture         =   "frmFtp.frx":172E
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2895
      Index           =   0
      Left            =   0
      TabIndex        =   3
      Top             =   1320
      Width           =   8415
      Begin VB.TextBox Text1 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1815
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   13
         Text            =   "frmFtp.frx":1A38
         Top             =   240
         Width           =   8175
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Nome File"
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
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   2280
         Width           =   8175
      End
   End
End
Attribute VB_Name = "frmFTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Automatico()
    
    cmdPulsante_Click (0)
    cmdPulsante_Click (6)
    cmdPulsante_Click (7)
    cmdPulsante_Click (1)
End Sub

Public Sub DownLoadFile()
    'Dichiarazione Variabili
    Dim bRet, bret1, bret2 As Boolean
    Dim szFileRemote As String, szDirRemote As String, szFileLocal As String
    Dim szTempString As String
    Dim nPos As Long, nTemp As Long
    Dim hFind As Long
    Dim findfile As WIN32_FIND_DATA
    Dim errore As Integer
    Dim count As Integer
    'Totale File Errati
    errore = 0
    'Totale File
    count = 0
    'Controllo della Connessione
    If bActiveSession Then
        'Imposto la directory Corrente
        Call FtpSetCurrentDirectory(hConnection, dirserv)
        'Inizio a Cercare i File
        hFind = FtpFindFirstFile(hConnection, "*.txt", findfile, 0, 0)
        If hFind = 0 Then
            Text1.Text = Text1.Text + "There aren't File for DownLoad ..." + vbCrLf
            Exit Sub
        End If
        count = 1
        szFileRemote = Trim(Mid(findfile.cFileName, 1, InStr(1, findfile.cFileName, Chr(0), vbTextCompare) - 1))
        Label1(1).Caption = szFileRemote
        Label1(1).Refresh
        bRet = FtpGetFile(hConnection, szFileRemote, dirrecv & "/" & szFileRemote, False, _
        INTERNET_FLAG_RELOAD, dwType, 0)
        If bRet = False Then
            'File Log'
            Text1.Text = Text1.Text + "Error: DownLoad File " + szFileRemote + " : " + CStr(Err.LastDllError) + vbCrLf
            Text1.Refresh
            Text1.SelStart = Len(Text1.Text)
            errore = errore + 1
        Else
            bret2 = FtpDeleteFile(hConnection, szFileRemote)
        End If
        'If bRet = False Then ErrorOut cstr(Err.LastDllError), "FtpGetFile"
        bret1 = InternetFindNextFile(hFind, findfile)
        While bret1 <> False
            szFileRemote = Trim(Mid(findfile.cFileName, 1, InStr(1, findfile.cFileName, Chr(0), vbTextCompare) - 1))
            count = count + 1
            Label1(1).Caption = szFileRemote
            Label1(1).Refresh
            bRet = FtpGetFile(hConnection, szFileRemote, dirrecv & "/" & szFileRemote, False, _
            INTERNET_FLAG_RELOAD, dwType, 0)
            If bRet = False Then
                'File Log'
                Text1.Text = Text1.Text + "Error DownLoad File " + szFileRemote + " : " + CStr(Err.LastDllError) + vbCrLf
                Text1.Refresh
                Text1.SelStart = Len(Text1.Text)
                errore = errore + 1
            Else
                bret2 = FtpDeleteFile(hConnection, szFileRemote)
            End If
            bret1 = InternetFindNextFile(hFind, findfile)
        Wend
        Label1(1).Caption = ""
        'File Log'
        If errore = 0 Then
            Text1.Text = Text1.Text + "DownLoad Complete ..." + vbCrLf
            Text1.Refresh
            Text1.SelStart = Len(Text1.Text)
        Else
            Text1.Text = Text1.Text + "DowLoad don't Complete ..." + vbCrLf
            Text1.Refresh
            Text1.SelStart = Len(Text1.Text)
        End If
        Text1.Text = Text1.Text + "Total file : " + CStr(count) + " Error File : " + CStr(errore) + vbCrLf
        Text1.Refresh
        Text1.SelStart = Len(Text1.Text)
    End If
End Sub
Private Sub DisconnettiServer()
    'Chiusura Connessione al Server
    If hConnection <> 0 Then InternetCloseHandle hConnection
    hConnection = 0
End Sub


Private Sub DisconnettiInternet()
    'Chiusura Connessione ad Internet
    If hOpen <> 0 Then InternetCloseHandle (hOpen)
    hOpen = 0
End Sub







Private Sub rcd(pszDir As String)
    '***********************
    'Trovo la directory principale
    '/ = Root
    '/Nomedirectory = SottoDirectory
    '***********************
    If pszDir <> "" Then
        Dim sPathFromRoot As String
        Dim bRet As Boolean
        If InStr(1, pszDir, server) Then
        sPathFromRoot = Mid(pszDir, Len(server) + 1, Len(pszDir) - Len(server))
        Else
        sPathFromRoot = pszDir
        End If
        If sPathFromRoot = "" Then sPathFromRoot = "/"
        bRet = FtpSetCurrentDirectory(hConnection, sPathFromRoot)
    End If
End Sub
Private Sub ScriviFileLog()
    
    Dim iFile As Integer
    iFile = FreeFile
    Open NomeFileLog For Append As #iFile
    Print #iFile, Text1.Text
    Close #iFile
    
End Sub

Private Sub UpLoad()
    Dim bRet As Boolean
    Dim szFileRemote As String, szDirRemote As String, szFileLocal As String
    Dim szTempString As String
    Dim szFileConPath As String
    Dim errore As Integer
    Dim count As Integer
    count = 0
    errore = 0
    If bActiveSession Then
        szTempString = server + dirserv
        szDirRemote = Right(szTempString, Len(szTempString) - Len(server))
        szFileLocal = Dir(dirsend + "*.txt")
        If Trim(szFileLocal) = "" Then
            Text1.Text = Text1.Text + "There aren't File for UpLoad  ..." + vbCrLf
            Text1.Refresh
            Text1.SelStart = Len(Text1.Text)
            Exit Sub
        End If
        
        While szFileLocal <> ""
            count = count + 1
            Label1(1).Caption = szFileLocal
            Label1(1).Refresh
            szFileConPath = dirsend + szFileLocal
            szFileRemote = szFileLocal
            If (szDirRemote = "") Then szDirRemote = "\"
            rcd szDirRemote
            bRet = FtpPutFile(hConnection, szFileConPath, szFileRemote, _
             dwType, 0)
            If bRet = False Then
                'File Log'
                Text1.Text = Text1.Text + "Error Upload File : " + CStr(Err.LastDllError) + vbCrLf
                Text1.Refresh
                Text1.SelStart = Len(Text1.Text)
                errore = errore + 1
            End If
            szFileLocal = Dir
        Wend
        'File Log'
        If errore = 0 Then
            Text1.Text = Text1.Text + "UpLoad File Complete ..." + vbCrLf
            Text1.Refresh
            Text1.SelStart = Len(Text1.Text)
        Else
            Text1.Text = Text1.Text + "UpLoad don't Complete ..." + vbCrLf
            Text1.Refresh
            
        End If
        Text1.Text = Text1.Text + "Total File : " + CStr(count) + " Error File : " + CStr(errore) + vbCrLf
        Text1.Refresh
        Text1.SelStart = Len(Text1.Text)
        Label1(1).Caption = ""
   End If
End Sub




Private Function ConnessioneInternet() As Boolean
    '***********************
    'Connessione ad internet
    '***********************
    hOpen = InternetOpen(scUserAgent, INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, 0)
    If hOpen = 0 Then
        'File Log'
        ConnessioneInternet = False
        Exit Function
    Else
        'File Log'
        ConnessioneInternet = True
        Exit Function
    End If
End Function

Private Function ConnessioneServer() As Boolean
    
    Dim nflag As Long
    If tipftp = 1 Then
        nflag = INTERNET_FLAG_PASSIVE
    Else
        nflag = 0
    End If
            
    hConnection = InternetConnect(hOpen, server, INTERNET_INVALID_PORT_NUMBER, _
    username, password, INTERNET_SERVICE_FTP, nflag, 0)
    If hConnection = 0 Then
        bActiveSession = False
        ConnessioneServer = False
        Exit Function
    Else
        bActiveSession = True
        ConnessioneServer = True
        Exit Function
    End If

End Function

Private Sub cmdPulsante_Click(Index As Integer)
    Select Case Index
        
        Case 0
            '***********************
            'Connetti
            '***********************
            Text1.Text = Text1.Text + "***************************************************************" + vbCrLf
            Text1.Text = Text1.Text + "Connect Day    : " + Format(Now, "dd/mm/yyyy") + vbCrLf
            Text1.Text = Text1.Text + "Hour           : " + Format(Now, "hh:mm:ss") + vbCrLf
            Text1.Text = Text1.Text + "User           : " + prgCol + vbCrLf
            Text1.Text = Text1.Text + "***************************************************************" + vbCrLf
            Text1.Text = Text1.Text + "Start Internet Connection ..." + vbCrLf
            Text1.Refresh
            Text1.SelStart = Len(Text1.Text)
            If Not ConnessioneInternet Then
                Text1.Text = Text1.Text + "Error Internet Connection : " + CStr(Err.LastDllError) + vbCrLf
                Text1.Refresh
                Text1.SelStart = Len(Text1.Text)
                Exit Sub
            Else
                Text1.Text = Text1.Text + "Internet Connection Complete ..." + vbCrLf
                Text1.Refresh
                Text1.SelStart = Len(Text1.Text)
                'File Log'
                Text1.Text = Text1.Text + "Start Server Connection ..." + vbCrLf
                Text1.Refresh
                Text1.SelStart = Len(Text1.Text)
                If Not ConnessioneServer Then
                    'File Log'
                    Text1.Text = Text1.Text + "Error : Server Connection : " + CStr(CStr(Err.LastDllError)) + vbCrLf
                    Text1.Refresh
                    Text1.SelStart = Len(Text1.Text)
                    Exit Sub
                Else
                    'File Log'
                    Text1.Text = Text1.Text + "Server Connection Complete ..." + vbCrLf
                    Text1.Refresh
                    Text1.SelStart = Len(Text1.Text)
                End If
                PreparaForm
            End If
        Case 1
            '***********************
            'Disconnetti
            '***********************
            DisconnettiServer
            DisconnettiInternet
            Text1.Text = Text1.Text + "Server Disconnect Complete ..." + vbCrLf
            Text1.Text = Text1.Text + "*************************Fine*******************************" + vbCrLf
            Text1.Refresh
            Text1.SelStart = Len(Text1.Text)
            ScriviFileLog
            bActiveSession = False
            PreparaForm
        Case 2
            '***********************
            'Setup INI
            '***********************
            Load frmSetup
            frmSetup.Show vbModal
        Case 3
            '*******************************************
            'Per salvataggio e Caricamento in automatico
            '*******************************************
            Automatico
            End
        Case 4
            '***********************
            'Uscita
            '***********************
            End
        Case 5
            '***********************
            'Apertura File Log
            '***********************
            Shell "notepad.exe " + NomeFileLog, vbNormalFocus
        Case 6
            '***********************
            'UpLoad
            '***********************
            PreparaForm "UpLoad"
            UpLoad
            PreparaForm "UpLoadFine"
        Case 7
            '***********************
            'DownLoad
            '***********************
            PreparaForm "DownLoad"
            DownLoadFile
            DisconnettiServer
            If ConnessioneServer Then
            End If
            PreparaForm "DownLoadFine"
    End Select
End Sub

