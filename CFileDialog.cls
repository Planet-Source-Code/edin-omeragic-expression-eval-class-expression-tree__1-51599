VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFileDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Api declarations
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetActiveWindow Lib "user32" () As Long
'Types
Private Type OPENFILENAME
        lStructSize As Long
        hwndOwner As Long
        hInstance As Long
        lpstrFilter As String
        lpstrCustomFilter As String
        nMaxCustFilter As Long
        nFilterIndex As Long
        lpstrFile As String
        nMaxFile As Long
        lpstrFileTitle As String
        nMaxFileTitle As Long
        lpstrInitialDir As String
        lpstrTitle As String
        flags As Long
        nFileOffset As Integer
        nFileExtension As Integer
        lpstrDefExt As String
        lCustData As Long
        lpfnHook As Long
        lpTemplateName As String
End Type

'Constants
Private Const OFN_ALLOWMULTISELECT = &H200
Private Const OFN_CREATEPROMPT = &H2000
Private Const OFN_ENABLEHOOK = &H20
Private Const OFN_ENABLETEMPLATE = &H40
Private Const OFN_ENABLETEMPLATEHANDLE = &H80
Private Const OFN_EXPLORER = &H80000
Private Const OFN_EXTENSIONDIFFERENT = &H400
Private Const OFN_FILEMUSTEXIST = &H1000
Private Const OFN_HIDEREADONLY = &H4
Private Const OFN_LONGNAMES = &H200000
Private Const OFN_NOCHANGEDIR = &H8
Private Const OFN_NODEREFERENCELINKS = &H100000
Private Const OFN_NOLONGNAMES = &H40000
Private Const OFN_NONETWORKBUTTON = &H20000
Private Const OFN_NOREADONLYRETURN = &H8000
Private Const OFN_NOTESTFILECREATE = &H10000
Private Const OFN_NOVALIDATE = &H100
Private Const OFN_OVERWRITEPROMPT = &H2
Private Const OFN_PATHMUSTEXIST = &H800
Private Const OFN_READONLY = &H1
Private Const OFN_SHAREAWARE = &H4000
Private Const OFN_SHAREFALLTHROUGH = 2
Private Const OFN_SHARENOWARN = 1
Private Const OFN_SHAREWARN = 0
Private Const OFN_SHOWHELP = &H10
Private Const OFS_MAXPATHNAME = 128



'Vars

Dim m_FullPath As String
Dim m_InitDir As String
Dim m_Filter As String
Dim m_FilterIndex As Long
Dim m_Title As String
Dim m_MultiSelect As Boolean
Dim m_FileNames() As String


'--------------FullPath-------------------
Property Get FullPath() As String
    FullPath = m_FullPath
End Property

Property Let FullPath(Path As String)
    m_FullPath = Path
End Property
'---------------InitDir-------------------
Property Get InitDir() As String
    InitDir = m_InitDir
End Property
Property Let InitDir(Value As String)
    m_InitDir = Value
End Property
'---------------DialogTitle---------------
Property Get DialogTitle() As String
    DialogTitle = m_Title
End Property
Property Let DialogTitle(Value As String)
    m_Title = Value
End Property
'---------------Filter--------------------
Property Get Filter() As String
    Filter = m_Filter
End Property
Property Let Filter(Value As String)
    m_Filter = Value
End Property
'---------------FilterIndex---------------
Property Get FilterIndex() As Long
    FilterIndex = m_FilterIndex
End Property
Property Let FilterIndex(Value As Long)
    m_FilterIndex = Value
End Property
'---------------MultiSelect---------------
Property Get MultiSelect() As Boolean
    MultiSelect = m_MultiSelect
End Property
Property Let MultiSelect(Value As Boolean)
    m_MultiSelect = Value
End Property

'--------------Count----------------------
Property Get Count() As Integer
    On Local Error Resume Next
    Count = UBound(m_FileNames)
    If Err.Number <> 0 Then
        Count = 0
    End If
End Property

'-------------FileName--------------------
Function FileName(Index As Integer) As String
    FileName = m_FileNames(Index)
End Function

Sub ClearNames()
    ReDim m_FileNames(0)
End Sub

Function Show(bOpen As Boolean, Optional hwnd As Long = 0) As Boolean
    
    If hwnd = 0 Then
        hwnd = GetActiveWindow
    End If
    
    Dim OFN As OPENFILENAME
    Dim RC  As Long
    ReDim m_FileNames(0)
    
    OFN.lStructSize = Len(OFN)
    OFN.hwndOwner = hwnd
    OFN.hInstance = OFN_EXPLORER

    OFN.lpstrFilter = Replace(m_Filter, "|", Chr(0))
        
    If m_MultiSelect = False Then
        OFN.lpstrFile = FullPath + String(260, vbNullChar)
    Else
        OFN.lpstrFile = FullPath + String(2600, vbNullChar)
    End If
    OFN.nMaxFile = Len(OFN.lpstrFile)
    
    OFN.lpstrFileTitle = String(260, vbNullChar)
    OFN.nMaxFileTitle = 260
    OFN.lpstrTitle = m_Title
    OFN.nFilterIndex = m_FilterIndex
    OFN.lpstrInitialDir = m_InitDir

    
    If bOpen Then
        If m_MultiSelect Then
            OFN.flags = OFN_ALLOWMULTISELECT
        End If
        OFN.flags = OFN.flags Or OFN_FILEMUSTEXIST
        RC = GetOpenFileName(OFN)
    Else
        OFN.flags = OFN.flags Or OFN_OVERWRITEPROMPT
        RC = GetSaveFileName(OFN)
    End If
    
    If RC = 0 Then GoTo ExitProc
    
    m_FullPath = TrimNullChar(OFN.lpstrFile)

    
    
    Dim I As Integer
    Dim J As Integer
    Dim K As Integer
    

    If m_MultiSelect And bOpen = True Then
        J = 1
        m_FullPath = m_FullPath
        Do
            I = InStr(J, m_FullPath, Chr(0))
            J = InStr(I + 1, m_FullPath, Chr(0))
            
            If I = 0 Then Exit Do
            If J = 0 Then J = Len(m_FullPath)
            ReDim Preserve m_FileNames(K)
            
            m_FileNames(K) = Mid(m_FullPath, I + 1, J - I - 1)

            K = K + 1
        Loop
        I = InStr(1, m_FullPath, Chr(0))
        m_FullPath = Left(m_FullPath, I - 1)
    End If
    
ExitProc:
    If RC <> 0 Then Show = True
End Function

Private Function TrimNullChar(str As String) As String
    Dim Pos As Integer
    Pos = InStr(1, str, Chr(0) + Chr(0)) - 1
    If Pos > 0 Then
        TrimNullChar = Left(str, Pos)
    End If
End Function


Private Sub Class_Initialize()
    m_Filter = "All files (*.*)|*.*||"
End Sub
