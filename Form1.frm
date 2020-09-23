VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Math, demo"
   ClientHeight    =   6765
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9870
   LinkTopic       =   "Form1"
   ScaleHeight     =   451
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   658
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Color"
      Height          =   405
      Left            =   2310
      TabIndex        =   20
      Top             =   2430
      Width           =   1125
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load"
      Height          =   375
      Left            =   1230
      TabIndex        =   19
      Top             =   6210
      Width           =   1125
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   90
      TabIndex        =   18
      Top             =   6210
      Width           =   1125
   End
   Begin VB.TextBox txtX 
      Height          =   375
      Left            =   330
      TabIndex        =   17
      Top             =   5280
      Width           =   1575
   End
   Begin VB.TextBox txtT2 
      Height          =   345
      Left            =   1680
      TabIndex        =   15
      Text            =   "6.3"
      Top             =   2010
      Width           =   735
   End
   Begin VB.TextBox txtT1 
      Height          =   345
      Left            =   960
      TabIndex        =   14
      Text            =   "0"
      Top             =   2010
      Width           =   645
   End
   Begin VB.TextBox txtInc 
      Height          =   315
      Left            =   960
      TabIndex        =   12
      Text            =   "0.1"
      Top             =   2400
      Width           =   645
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Add"
      Height          =   345
      Left            =   2850
      TabIndex        =   10
      Top             =   420
      Width           =   585
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      Height          =   315
      Left            =   2880
      TabIndex        =   7
      Top             =   1530
      Width           =   585
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Draw"
      Height          =   375
      Left            =   2340
      TabIndex        =   5
      Top             =   5220
      Width           =   1125
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000008&
      Height          =   5295
      Left            =   3600
      ScaleHeight     =   9.234
      ScaleMode       =   7  'Centimeter
      ScaleWidth      =   10.821
      TabIndex        =   2
      Top             =   120
      Width           =   6195
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Draw tangent"
      Height          =   375
      Left            =   300
      TabIndex        =   1
      Top             =   5700
      Width           =   1605
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Text            =   "pow(x,2)"
      Top             =   450
      Width           =   2685
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Text            =   "sin(t)*4"
      Top             =   1530
      Width           =   2715
   End
   Begin VB.TextBox Text3 
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Text            =   "cos(t)*4"
      Top             =   1200
      Width           =   2715
   End
   Begin VB.ListBox List1 
      Height          =   2310
      ItemData        =   "Form1.frx":0000
      Left            =   120
      List            =   "Form1.frx":0002
      Style           =   1  'Checkbox
      TabIndex        =   6
      Top             =   2880
      Width           =   3345
   End
   Begin VB.Label Label5 
      Caption         =   "x"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   5340
      Width           =   195
   End
   Begin VB.Label Label4 
      Caption         =   "Range"
      Height          =   195
      Left            =   120
      TabIndex        =   13
      Top             =   2100
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Increment"
      Height          =   315
      Left            =   120
      TabIndex        =   11
      Top             =   2430
      Width           =   825
   End
   Begin VB.Label Label2 
      Caption         =   "y(x)"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   180
      Width           =   465
   End
   Begin VB.Label Label1 
      Caption         =   "x(t) , y(t)"
      Height          =   225
      Left            =   120
      TabIndex        =   8
      Top             =   930
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdLoad_Click()
    Dim d As New CFileDialog
    Dim fn As String, K As Integer
    d.Filter = "Expresions *.exps|*.exps"

    Dim exp As String
    Dim Color As String
    Dim sel As Boolean
    If d.Show(True) Then
        List1.Clear
        fn = d.FullPath
        Open fn For Input As #1
            Do Until EOF(1)
                Input #1, exp, Color, sel
                
                List1.AddItem exp
                List1.ItemData(List1.NewIndex) = Color
                List1.Selected(List1.NewIndex) = sel
            Loop
        Close 1
    End If
End Sub

Private Sub cmdSave_Click()
    Dim d As New CFileDialog
    Dim fn As String, K As Integer
    d.Filter = "Expresions *.exps|*.exps"
    If d.Show(False) Then
        fn = d.FullPath
        Open fn For Output As #1
            For K = 0 To List1.ListCount - 1
                Write #1, List1.List(K), List1.ItemData(K), List1.Selected(K)
            Next
        Close 1
    End If
End Sub

Private Sub Command1_Click()
    Dim s As String
    
    s = Text3.Text + " ; " + Text2.Text
    If Len(Trim(s)) > 2 Then
        If Not ItemExists(s) Then
            List1.AddItem s
            List1.ItemData(List1.NewIndex) = vbWhite
            List1.Selected(List1.NewIndex) = True
        End If
    End If
    
End Sub

Private Sub Command2_Click()
If Text1.Text <> "" Then
    If Not ItemExists(Text1.Text) Then
        List1.AddItem Text1.Text
        List1.ItemData(List1.NewIndex) = vbWhite
        List1.Selected(List1.NewIndex) = True
    End If
End If
End Sub

Function ItemExists(s As String) As Boolean
    Dim I As Integer
    For I = 0 To List1.ListCount - 1
        If Trim(LCase(List1.List(I))) = Trim(LCase(s)) Then
            ItemExists = True
            Exit Function
        End If
    Next

End Function


Private Sub Command3_Click()
If List1.ListIndex < 0 Then Exit Sub
    Dim c As New CColorDialog
    c.hwnd = Me.hwnd
    If c.Show Then
        List1.ItemData(List1.ListIndex) = c.Color
    End If
End Sub

Private Sub Command4_Click()
    On Error Resume Next
    
    Command6_Click
    
    Dim exp As New ExpTree, exp1 As New ExpTree
    Dim sExp As String
    Dim x As Double, Pos As Double
    x = Val(txtX)
    
    sExp = List1.Text
    Pos = InStr(1, sExp, ";")
    
    If Pos = 0 Then
        exp.Compile sExp
    Else
        exp.Compile Left(sExp, Pos - 1)
        exp1.Compile Mid(sExp, Pos + 1)
    End If
    
    Dim sw  As Single, sh As Single
    sw = Picture1.ScaleWidth
    sh = Picture1.ScaleHeight
    
    
    Dim y1 As Double
    Dim y2 As Double
    Dim K As Double
    Dim n As Double
    Dim v1 As Double, v2 As Double
    
    If Pos = 0 Then
        exp.SetVar "x", x
        y1 = exp.Value
        exp.SetVar "x", x + 0.0001
        y2 = exp.Value
        K = (y2 - y1) / 0.0001
        n = y1 - x * K
      'y() = kx+n
      'y(x-3) = k(x-3) + n
      'y(x+3) = k(x+3) + n
        v1 = (x - 3) * K + n
        v2 = (x + 3) * K + n
    Else
'        MsgBox "Not supported.", vbInformation
'        exp.SetVar "t", x
'        y1 = exp.Value
'        exp.SetVar "t", x + 0.0001
'        y2 = exp.Value
'
'        exp1.SetVar "t", x
'        v1 = exp.Value
'        exp1.SetVar "t", x + 0.0001
'        v2 = exp.Value
'
'        k = (y2 - y1) / (v2 - v1)
'        exp1.SetVar "t", x
'        exp.SetVar "t", exp1.Value
'        y1 = exp.Value
'        n = y1 - k * x
'
'        'k = dy/dx = dy/dt * dx/dt
        Exit Sub
    End If
    Picture1.ForeColor = vbYellow
    Picture1.Circle (sw / 2 + x, sh / 2 - y1), 0.1
    Picture1.Line (sw / 2 + x - 3, sh / 2 - v1)-(sw / 2 + x + 3, sh / 2 - v2)
End Sub

Private Sub Command6_Click()
    On Error Resume Next
    Static bWork As Boolean
    
    If bWork Then Exit Sub
    
    bWork = True
    
    Picture1.Cls
    DrawAxis Picture1
    
    Dim tm As Single: tm = Timer
    Dim dt As Double, t1 As Double, t2 As Double
    Dim t As Double, sh2 As Double, sw2 As Double
    
    dt = Val(txtInc.Text)
    If dt <= 0 Then dt = 0.1
    
    t1 = Val(txtT1.Text)
    t2 = Val(txtT2.Text)
    
    If t2 <= t1 Then
        t1 = 0
        t2 = 6.3
    End If
    

    Dim sw  As Single, sh As Single
    sw = Picture1.ScaleWidth
    sh = Picture1.ScaleHeight
    sw2 = sw / 2
    sh2 = sh / 2
    
    Dim exp1 As New ExpTree
    Dim exp2 As New ExpTree
    exp1.StrVars = "t|x"
    exp2.StrVars = "t|x"
    
    Dim K As Integer
    Dim sExp As String
    Dim sExp1 As String, sExp2 As String, Pos As String
    
    For K = 0 To List1.ListCount - 1
        If Not List1.Selected(K) Then GoTo EndFor
        sExp = List1.List(K)
        Pos = InStr(1, sExp, ";")
        If Pos > 0 Then
            sExp1 = Left(sExp, Pos - 1)
            sExp2 = Mid(sExp, Pos + 1)
            'draw
            exp1.Compile sExp1
            exp2.Compile sExp2
            
            exp1.SetVar "t", t1
            exp2.SetVar "t", t1
            
            Picture1.ForeColor = List1.ItemData(K)
            Picture1.PSet (sw2 + exp1.Value, sh2 - exp2.Value)
            For t = t1 To t2 Step dt
                exp1.SetVar "t", t
                exp2.SetVar "t", t
                Picture1.Line -(sw2 + exp1.Value, sh2 - exp2.Value)
                DoEvents
            Next
        Else
            exp1.Compile sExp
            exp1.SetVar "x", -sw2
            Picture1.ForeColor = List1.ItemData(K)
            Picture1.PSet (0, sh2 - exp1.Value)
            For t = -sw2 To sw2 Step dt
                exp1.SetVar "x", t
                Picture1.Line -(sw2 + t, sh2 - exp1.Value)
                DoEvents
            Next
        End If
EndFor:
    Next
    Caption = "time:" + CStr(Round(Timer - tm, 5))
    bWork = False
End Sub

Private Sub Form_Resize()
On Error Resume Next
    Picture1.Width = Me.ScaleWidth - Picture1.Left - 5
    Picture1.Height = Me.ScaleHeight - Picture1.Top - 5
End Sub

Sub DrawAxis(p As PictureBox)
    Dim sw  As Single, sh As Single
    sw = p.ScaleWidth
    sh = p.ScaleHeight
    p.ForeColor = vbWhite
    p.Line (0.3, sh / 2)-(sw - 0.3, sh / 2)
    p.Line (sw / 2, 0.3)-(sw / 2, sh - 0.3)
    
    p.Line (sw - 0.3, sh / 2)-(sw - 0.6, sh / 2 - 0.1)
    p.Line (sw - 0.3, sh / 2)-(sw - 0.6, sh / 2 + 0.1)
    'Picture1.Line (sw - 0.6, sh / 2 - 0.1)-(sw - 0.6, sh / 2 + 0.1)
    p.Line (sw / 2, 0.3)-(sw / 2 - 0.1, 0.6)
    p.Line (sw / 2, 0.3)-(sw / 2 + 0.1, 0.6)
    Dim x As Double
    
    For x = sw / 2 To sw - 0.6 Step 1
        p.Line (x, sh / 2 - 0.1)-(x, sh / 2 + 0.1)
    Next
    For x = sh / 2 To 0.6 Step -1
        p.Line (sw / 2 - 0.1, x)-(sw / 2 + 0.1, x)
    Next
    For x = sw / 2 To 0.6 Step -1
        p.Line (x, sh / 2 - 0.1)-(x, sh / 2 + 0.1)
    Next
    For x = sh / 2 To sh - 0.6 Step 1
        p.Line (sw / 2 - 0.1, x)-(sw / 2 + 0.1, x)
    Next

End Sub
Private Sub List1_KeyPress(KeyAscii As Integer)
    If List1.ListIndex < 0 Then Exit Sub
    If KeyAscii = Asc("d") Or KeyAscii = Asc("D") Then
    
        List1.RemoveItem List1.ListIndex
    End If
End Sub


Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        txtX = Replace(CStr(Round(x - Picture1.ScaleWidth / 2, 3)), ",", ".")
    End If
End Sub

Private Sub picture1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    txtX = Replace(CStr(Round(x - Picture1.ScaleWidth / 2, 3)), ",", ".")
End Sub

Private Sub Picture1_Paint()
If Val(txtInc.Text) > 0.09 Then
    Command6_Click
Else
    DrawAxis Picture1
End If
End Sub

Private Sub Picture1_Resize()
    Picture1.Refresh
End Sub
