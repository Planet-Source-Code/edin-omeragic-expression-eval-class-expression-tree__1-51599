VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExpNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=======================================
'|autor: Edin Omeragic
'|datum: 03.04.2004 , 16:00 sati
'|comment: napravljeno od nule (build from empty project)
'=======================================
'|thanks^2 to David J. Eck
'|for tips and super-nice book "Introduction to Programming Using Java"
'|link: http://math.hws.edu/javanotes
'=======================================

'about coding: This is vb, and it wouldnot be nice to use several clases
'              to respresent diferent kind of nodes, insted this i used
'              variable Kind ( as NodeKind) to indentify kind of node ;)
'              Another thing, calculating woud be faster if "op" variable
'              holds id of operation like
'              ("+" => 1, "-" => 2 ..., "sin" => 101, "cos" => 102
'              since its fast enough i didnt bother
'              >>Errors are ignored - function returns zero
'              TO DO: Build class of user defined functions
'              TO DO: Build class of constants (like g=9.81, e, pi)
'              To DO: Conect classes with this one

'NOTE:         >>since its easier to give angles in degrees use operator @
'              Operator @  converts degrees to radians
'              sin(45@) is eqvivalent to: sin(pi/4)
Option Explicit
Private Const pi As Double = 3.14159265358979
Private Const pi2 As Double = pi / 2
Private Const e As Double = 2.718281828

Public Enum NodeKind
    NodeValue
    NodeOp
    NodeFunction
    NodeVariable
End Enum

Public sExp As String
Public Kind As NodeKind
Public LeftNode As ExpNode
Public RightNode As ExpNode
Public m_Value As Double
Public op As String 'operator , function or variable name


Private Sub Class_Initialize()
    Kind = NodeValue
End Sub
Function SetFunc(fname, L As ExpNode, R As ExpNode)
    Kind = NodeFunction
    op = fname
    Set LeftNode = L
    Set RightNode = R
End Function

Sub SetOp(op_code As String, L As ExpNode, R As ExpNode)
    Kind = NodeOp
    op = op_code
    Set Me.LeftNode = L
    Set Me.RightNode = R
End Sub

Sub SetValue(ByVal v As Double)
    Kind = NodeValue
    m_Value = v
End Sub
Function GetValue() As Double
On Error GoTo ERR_LINE
    Dim x As Double, y As Double, t As Double, t1 As Double
    
    If Kind = NodeValue Then
        GetValue = m_Value
    ElseIf Kind = NodeVariable Then
        'if this is first variable then value it contines value
        If LeftNode Is Nothing Then
            GetValue = m_Value
        Else
            'left node points to first variable found in the expresion
            GetValue = LeftNode.GetValue
            'left node is pointer to first variable
            'after this line above is executed, called from other instance of
            'ExpNode
        End If
    ElseIf Kind = NodeOp Then
        Select Case op
            Case "+": GetValue = LeftNode.GetValue + RightNode.GetValue
            Case "-": GetValue = LeftNode.GetValue - RightNode.GetValue
            Case "*": GetValue = LeftNode.GetValue * RightNode.GetValue
            Case "/":
                y = RightNode.GetValue
                If y <> 0 Then
                    GetValue = LeftNode.GetValue / y
                Else
                    'error: Division by zero
                    GetValue = 0
                End If
            Case "^"
                x = LeftNode.GetValue
                y = RightNode.GetValue
                If x <> 0 And y <> 0 Then
                    GetValue = x ^ y
                Else
                    GetValue = 0
                End If
            Case "@"
                'converts angles to radians
                GetValue = LeftNode.GetValue * pi / 180
            Case "!"
                'factoriel
                GetValue = Fact(LeftNode.GetValue)
        End Select
    ElseIf Kind = NodeFunction Then
        If LeftNode Is Nothing = False Then
            x = LeftNode.GetValue
        End If
        If RightNode Is Nothing = False Then
            y = RightNode.GetValue
        End If
        
        Select Case LCase(op)
            Case "sin": GetValue = Sin(x)
            Case "cos": GetValue = Cos(x)
            Case "tan": GetValue = Tan(x)
            Case "abs": GetValue = Abs(x)
            Case "exp": GetValue = exp(x)
            Case "log": If x > 0 Then GetValue = Log(x)
            Case "ln":  If x > 0 Then GetValue = Log(x) / Log(e)
            Case "rnd": GetValue = Rnd()
            Case "fix": GetValue = Fix(x)
            Case "int": GetValue = Int(x)
            Case "abs": GetValue = Abs(x)
            Case "sqr": GetValue = Sqr(x)
            Case "sgn": GetValue = Sgn(x)
            Case "asin":
                If x = 1 Then
                    GetValue = pi2
                ElseIf x = -1 Then
                    GetValue = -pi2
                ElseIf x > -1 And x < 1 Then
                    GetValue = Atn(x / Sqr(1 - x ^ 2))
                Else
                    'set error flag
                End If
            Case "acos"
                If x = 1 Then
                    GetValue = 0
                ElseIf x = -1 Then
                    GetValue = pi
                ElseIf x > -1 And x < 1 Then
                    GetValue = Atn(-x / Sqr(1 - x ^ 2)) + pi2
                Else
                    'set error flag
                End If
            Case "asec": GetValue = pi2 - Atn(Sgn(x) / Sqr(x ^ 2 - 1))
            Case "acsc": GetValue = Atn(Sgn(x) / Sqr(x ^ 2 - 1))
            Case "acot": GetValue = pi2 - Atn(x)
            Case "sinh": GetValue = (e ^ x - e ^ (-x)) / 2
            Case "cosh": GetValue = (e ^ x + e ^ (-x)) / 2
            Case "tanh"
                t = e ^ x: t1 = e ^ (-x)
                GetValue = (t - t1) / (t + t1)
            Case "sech": GetValue = 2 / (e ^ x + e ^ (-x))
            Case "csch": GetValue = 2 / (e ^ x - e ^ (-x))
            Case "coth":
                t = e ^ x: t1 = e ^ (-x)
                GetValue = (t + t1) / (t - t1)
            Case "asinh": GetValue = Log(x + Sqr(x ^ 2 + 1))
            Case "acosh": GetValue = Log(x + Sqr(x ^ 2 - 1))
            Case "atanh": GetValue = Log((1 + x) / (1 - x)) / 2
            Case "asech": GetValue = Log(Sqr(1 - x ^ 2) / x)
            Case "acsch": GetValue = Log(Sgn(x) * Sqr(1 - x ^ 2) / x)
            Case "acoth": GetValue = Log((1 + x) / (x - 1)) / 2
            Case "sec"
                t = Cos(x): If t <> 0 Then GetValue = 1 / t
            Case "csc"
                t = Sin(x): If t <> 0 Then GetValue = 1 / t
            Case "ctan"
                t = Tan(x): If t <> 0 Then GetValue = 1 / t
            
            '===============================================
            'functions of two arguments
            '===============================================
            Case "pow"
                GetValue = x ^ y
            Case "root"
                GetValue = x ^ (1 / y)
            Case "above"
                If x > y Then GetValue = 1 Else GetValue = 0

            Case Else
                'scan for user defined functions
                '(simply set of ExpTree s)
                'Debug.Print "Unsuported function:"; op
        End Select
    End If
    Exit Function
ERR_LINE:
End Function
Private Function Fact(ByVal x As Double) As Double
    On Error Resume Next
    Dim i As Long
    Dim f As Double: f = 1
    i = Fix(x)
    If i < 1 Or i > 170 Then
        Fact = 0
        Exit Function
        'set error flag
    End If
    Select Case i
        Case Is < 21
            Select Case i
                Case Is = 1: f = 1
                Case Is = 2: f = 2
                Case Is = 3: f = 6
                Case Is = 4: f = 24
                Case Is = 5: f = 120
                Case Is = 6: f = 720
                Case Is = 7: f = 5040
                Case Is = 8: f = 40320
                Case Is = 9: f = 362880
                Case Is = 10: f = 3628800
                Case Is = 11: f = 39916800
                Case Is = 12: f = 479001600
                Case Is = 13: f = 6227020800#
                Case Is = 14: f = 87178291200#
                Case Is = 15: f = 1307674368000#
                Case Is = 16: f = 20922789888000#
                Case Is = 17: f = 355687428096000#
                Case Is = 18: f = 6.402373705728E+15
                Case Is = 19: f = 1.21645100408832E+17
                Case Is = 20: f = 2.43290200817664E+18
            End Select
        Case Else
            f = 2.43290200817664E+18
            For i = 21 To x
                f = f * i
            Next
    End Select
    
    Fact = f
End Function

Private Sub Class_Terminate()
    Set LeftNode = Nothing
    Set RightNode = Nothing
    'Debug.Print "destroy:"; sExp
End Sub
