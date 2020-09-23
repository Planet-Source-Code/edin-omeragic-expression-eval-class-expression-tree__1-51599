VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CColorDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_Hwnd As Long
Private m_Colors(0 To 15) As Long
Private m_Color As Long
Private Type TCHOOSECOLOR
        lStructSize As Long
        hwndOwner As Long
        hInstance As Long
        rgbResult As Long
        lpCustColors As Long
        flags As Long
        lCustData As Long
        lpfnHook As Long
        lpTemplateName As String
End Type
Private Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As TCHOOSECOLOR) As Long
Function Show() As Boolean
    Dim TC As TCHOOSECOLOR
    Dim RC As Long
    With TC
        .hwndOwner = m_Hwnd
        .lStructSize = Len(TC)
        .rgbResult = 0
        .lpCustColors = VarPtr(m_Colors(0))
        .flags = CC_ANYCOLOR Or CC_RGBINIT
    End With
    RC = ChooseColor(TC)
    
    If RC = 1 Then
        m_Color = TC.rgbResult
        Show = True
    End If
End Function

Public Property Get hWnd() As Long
    hWnd = m_Hwnd
End Property

Public Property Let hWnd(ByVal vNewValue As Long)
    m_Hwnd = vNewValue
End Property

Private Sub Class_Initialize()
    For i = 0 To 15
        m_Colors(i) = vbWhite
    Next
End Sub

Public Property Get Color() As Long
    Color = m_Color
End Property


