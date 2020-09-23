VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExpTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=======================================
'|autor: Edin Omeragic
'|datum: 03.04.2004 , 16:00 sati
'|comment: napravljeno od nule
'=======================================
'|thanks^2 to David J. Eck
'|for tips and super-nice book "Introduction to Programming Using Java"
'|link: http://math.hws.edu/javanotes
'=======================================
Option Explicit
Public Root As ExpNode
Public Vars As New Collection
Public StrVars  As String
Public StoredValue As Double

Function SetVar(name As String, ByVal Value As Double) As Boolean
    Dim itm As ExpNode
    If VarExists(name) Then
        Set itm = Vars(name)
        itm.m_Value = Value ' change value for node variable
        SetVar = True
    End If
End Function

Function VarExists(name As String) As Boolean
    Dim itm As ExpNode
    On Error Resume Next
    Set itm = Vars(name)
    If Err.Number = 0 Then
        VarExists = True
        Err.Clear
    End If
End Function

Property Get Value() As Double
    If Root Is Nothing = False Then
        StoredValue = Root.GetValue
        Value = StoredValue
    End If
End Property
Sub PrintValue()
    Debug.Print "Value is:"; Value
End Sub

Sub Compile(ByVal sExp As String)
    Set Root = Nothing
    Set Vars = Nothing
    Set Vars = New Collection
    Set Root = New ExpNode
    
    sExp = Replace(sExp, " ", "")
    sExp = RemoveBrackets(sExp)
    sExp = Replace(sExp, "-", "+-")
    sExp = Replace(sExp, "++", "+")
    sExp = Replace(sExp, ")(", ")*(")
    'without this cant parese -a-a
    'because  0-(a-a) = 0 -0 = 0
    'whit +- becomes: +-a+-a => (0+ (-a)) + (-a) = (-a) - a
    Root.sExp = sExp
    Parse Root
'    Stop
End Sub

Public Sub Parse(n As ExpNode)
   ' Debug.Print "Parse:"; n.sExp
    Dim op As String
    Dim pos As Integer
    Dim arg As String, arg2 As String
    
    Dim tExp As String
    tExp = Trim(n.sExp)
    
    If IsNumeric(tExp) Then
        n.SetValue (Val(tExp))
        Exit Sub
    ElseIf Left(tExp, 1) = "-" Then
        'sign change 'eg. "-(1+1)" => "0-(1+1)"
        'Debug.Print "sign change:"; tExp
        tExp = "0" + tExp
        n.sExp = tExp
    ElseIf tExp = "" Then
        'for empty expresion
        n.SetValue 0
        Exit Sub
    ElseIf IsVar(tExp) Then
        n.Kind = NodeVariable
        n.op = tExp ' name of variable, but its not important
        n.m_Value = 0 ' initial value
        tExp = LCase(tExp)
        If Not VarExists(tExp) Then
            Vars.Add n, tExp ' if there is no link to variable in collection class
        Else
            'eg  expresion like  x^2 + x
            'Colection vars holds one reference to x (right one)
            'parsing x^2 creates new variable node, to make it simple i
            'just link this node with the one from collction
            'Method GetValue of node returns m_Value of Linked node
            'Why? - When you update variable its  enough to update
            'only one node in collection
            Set n.LeftNode = Vars(tExp)
            'variables can be located on more than on place in expresion
            'link left unused node to node with the same variable
            'other part is done in ExpNode
        End If
        Exit Sub
    End If
    
    'low priority operators are scaned first
    'FindOp and FindFunc are core of parsing routine
    pos = FindOp(n.sExp, "+", op)
    If pos = 0 Then
        pos = FindOp(n.sExp, "-", op)
    End If
    If pos = 0 Then
        pos = FindOp(n.sExp, "*/", op)
    End If
    If pos = 0 Then
        pos = FindOp(n.sExp, "^", op)
    End If
    If pos = 0 Then
        pos = FindOp(n.sExp, "@!", op)
    End If
    If pos = 0 Then
        op = FindFunc(n.sExp, arg, arg2)
    End If
    
    If Len(op) = 1 Then
        'operator
        n.Kind = NodeOp
        n.op = op
        If op <> "@" Or op <> "!" Then
            Set n.LeftNode = New ExpNode
            Set n.RightNode = New ExpNode
            n.LeftNode.sExp = RemoveBrackets(ExpLeft(n.sExp, pos))
            n.RightNode.sExp = RemoveBrackets(ExpRight(n.sExp, pos))
        Else
            Set n.LeftNode = New ExpNode
            Set n.RightNode = Nothing
            n.LeftNode.sExp = RemoveBrackets(ExpLeft(n.sExp, pos))
        End If
    ElseIf Len(op) > 1 Then
        'function
        n.Kind = NodeFunction
        n.op = op
        Set n.LeftNode = New ExpNode
        n.LeftNode.sExp = RemoveBrackets(arg)
        If arg2 <> "" Then
            Set n.RightNode = New ExpNode
            n.RightNode.sExp = RemoveBrackets(arg2)
        End If
        'Debug.Print "Arg:"; arg
        'function
    End If
    If n.LeftNode Is Nothing = False Then
        Parse n.LeftNode
    End If
    If n.RightNode Is Nothing = False Then
        Parse n.RightNode
    End If
End Sub
Private Function RemoveBrackets(ByVal sExp As String) As String
    '"(" "(" ... <expresion> ... ")"  ")" => <expresion>
    Dim change As Boolean
    Dim tExp As String
    Do
        change = False
        sExp = Trim(sExp)
        If Left(sExp, 1) = "(" Then
            If Right(sExp, 1) = ")" Then
                tExp = Mid(sExp, 2, Len(sExp) - 2)
                If ValidRemove(tExp) Then
                    sExp = tExp
                    change = True
                End If
            End If
        End If
    Loop While change
    RemoveBrackets = sExp
End Function
Private Function ValidRemove(sExp As String) As Boolean
    Dim p1 As Integer
    Dim p2 As Integer
    p1 = InStr(1, sExp, ")")
    ValidRemove = False
    If p1 > 0 Then
        p2 = InStr(1, Left(sExp, p1), "(")
        If p2 > 0 Then
            ValidRemove = True
            Debug.Print "can remove:"; sExp
        End If
    Else
        ValidRemove = True
    End If
End Function

Private Function ExpLeft(sExp As String, pos As Integer) As String
    If pos > 1 Then
        ExpLeft = Left(sExp, pos - 1)
    End If
End Function
Private Function ExpRight(sExp As String, pos As Integer) As String
    If pos < Len(sExp) Then
        ExpRight = Mid(sExp, pos + 1)
    End If
End Function

Public Function FindFunc(ByVal sExp As String, sArg As String, Optional sArg2 As String) As String
    ' <function> "(" <expresion> ")"
    Dim pos1 As Integer
    Dim pos2 As Integer
    Dim pos3 As Integer
    
    Dim func As String
    
    pos1 = InStr(1, sExp, "(")
    pos2 = InStrRev(sExp, ")")
    Dim xx As String: xx = sExp
    Dim op As String
    'lasy to make this right
    If pos1 > 0 Then Mid(xx, pos1, 1) = " "
    If pos2 > 0 Then Mid(xx, pos2, 1) = " "
    
    pos3 = FindOp(xx, ",", op) 'it's not operator but this
                               'function can be used for things like that
    If pos1 < 1 Or pos2 < pos1 Then
        'variable found or number
        Exit Function
    End If
    
    func = Left(sExp, pos1 - 1)
    
    FindFunc = Trim(func)
    
    Dim sLen As Integer
    sLen = pos2 - pos1 - 1
    If sLen > 0 Then
        If pos3 = 0 Then
            sArg = Mid(sExp, pos1 + 1, sLen)
        Else
            sArg = Mid(sExp, pos1 + 1, pos3 - pos1 - 1)
            sArg2 = Mid(sExp, pos3 + 1, pos2 - pos3 - 1)
        End If
    End If
End Function

Private Function FindOp(sExp As String, op_list As String, found_op As String) As Integer
    Dim i As Integer
    Dim op As String
    Dim pos As Integer
    Dim ccr As Integer
    Dim ccl As Integer
    For i = 1 To Len(op_list)
        op = Mid(op_list, i, 1)
        pos = 1
        Do
            pos = InStr(pos + 1, sExp, op)
            If pos = 0 Then Exit Do
            ccr = CharCount(sExp, pos, "(")
            ccl = CharCount(sExp, pos, ")")
            If ccr = ccl Then
                FindOp = pos
                found_op = op
                Exit Function
            End If
        Loop While pos > 0
    Next
    FindOp = 0
End Function
 
Private Function CharCount(ByVal sExp As String, ByVal sLen As Integer, ByVal ch As String) As Integer
    Dim i As Integer
    Dim tch As String
    Dim c As Integer: c = 0
    If sLen > Len(sExp) Then sLen = Len(sExp)
    For i = 1 To sLen
        tch = Mid(sExp, i, 1)
        If tch = ch Then
            c = c + 1
        End If
    Next
    CharCount = c
End Function

Private Function FuncExists(ByVal func As String) As Boolean
    'for error checking - Not implemented yet
    Const f  As String = "sin|cos|tan|ctg|abs|log|ln| etc."
    func = LCase(Trim(func))
    FuncExists = (InStr(1, f, func) > 0)
End Function

Function IsVar(s As String) As Boolean
    IsVar = InStr(1, StrVars, LCase(s))
End Function

Private Sub Class_Initialize()
    'names of variables used in expresion
    StrVars = "x|y|t|u|v"
End Sub
