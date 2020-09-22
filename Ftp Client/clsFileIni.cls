VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFileIni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

#If Win16 Then
  Private Declare Function WritePrivateProfileString Lib "KERNEL" (ByVal AppName As String, ByVal KeyName As String, ByVal keydefault As String, ByVal FileName As String) As Integer
  Private Declare Function GetPrivateProfileString Lib "KERNEL" (ByVal AppName As String, ByVal KeyName As String, ByVal keydefault As String, ByVal ReturnString As String, ByVal NumBytes As Integer, ByVal FileName As String) As Integer
#ElseIf Win32 Then
  Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal AppName As String, ByVal KeyName As String, ByVal keydefault As String, ByVal FileName As String) As Long
  Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal AppName As String, ByVal KeyName As String, ByVal keydefault As String, ByVal ReturnString As String, ByVal NumBytes As Long, ByVal FileName As String) As Long
#End If

Private m_strFileINI As String
Private m_strSection As String
Private m_strKey As String
Private m_strDescription As String

'Costanti
Private Const BufSize = 255


Public Property Let FileINI(strFileIni As String)

    m_strFileINI = strFileIni

End Property

Public Property Get FileINI() As String

    FileINI = m_strFileINI

End Property
Public Property Let Section(strSection As String)

    m_strSection = strSection

End Property

Public Property Get Section() As String

    Section = m_strSection

End Property

Public Property Let Key(strKey As String)

    m_strKey = strKey

End Property

Public Property Get Key() As String

    Key = m_strKey

End Property

Public Property Let Description(strDescription As String)

    m_strDescription = strDescription

End Property

Public Property Get Description() As String

    Description = m_strDescription

End Property


Public Sub AddToINI()
Dim rc As Long
    
    rc = WritePrivateProfileString(m_strSection, m_strKey, m_strDescription, m_strFileINI)
    
End Sub

Public Function ReadFromINI() As String
Dim rc As Long
Dim RetStr As String

    RetStr = Space(BufSize)
    rc = GetPrivateProfileString(m_strSection, m_strKey, "Non Trovato", RetStr, BufSize, m_strFileINI)
    
    
    RetStr = Left(RetStr, rc)
    If RetStr = "Non Trovato" Then
          m_strDescription = "Non Trovato"
       Else
          m_strDescription = RetStr
    End If
    
    ReadFromINI = m_strDescription
    
End Function


