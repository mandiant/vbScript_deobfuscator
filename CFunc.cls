VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFunc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public name As String
Public args As Collection
Private m_body() As String
Private n_body() As String 'newBody
Public ft As funcTypes
Public rawProto As String
Public LineNumber As Long
Public LastLine As String 'standardized

Public newProto As String
Public newName As String

Public localVars As New Collection

Public Enum funcTypes
    ft_notset = 0
    ft_function
    ft_sub
    ft_property
End Enum

Property Get gsMarker() As String
    gsMarker = funcOpenMarker & " func " & Me.name & " --------]"
End Property

Property Get body(Optional complete As Boolean = False) As String
    If complete Then
        body = rawProto & vbCrLf & Join(m_body, vbCrLf) & vbCrLf & LastLine
    Else
        body = Join(m_body, vbCrLf)
    End If
End Property

Property Get newBody(Optional complete As Boolean = True) As String
    If complete Then
        newBody = newProto & vbCrLf & Join(n_body, vbCrLf) & vbCrLf & LastLine
    Else
        newBody = Join(n_body, vbCrLf)
    End If
End Property



Sub SetFirstLine(X, lNo)
    
    rawProto = X
    LineNumber = lNo
    Set args = New Collection
    
    a = Replace(X, "private", Empty, , 1, vbTextCompare)
    a = Replace(a, "public", Empty, , 1, vbTextCompare)
    a = Replace(a, "friend", Empty, , 1, vbTextCompare)
    a = Replace(a, "sub", Empty, , 1, vbTextCompare)
    a = Replace(a, "function", Empty, , 1, vbTextCompare)
    a = Replace(a, vbTab, Empty)
    b = InStr(1, a, "(")
    If b > 2 Then
        name = Mid(a, 1, b - 1)
        ParseArgs Mid(a, b)
    Else
        name = a
    End If
    
    name = Trim(name)
    
End Sub

Private Sub ParseArgs(ByVal argString)
    
    
    argString = Replace(argString, vbTab, Empty)
    
    If Right(argString, 1) = ")" Then
        argString = Mid(argString, 1, Len(argString) - 1)
    End If
    
    If Left(argString, 1) = "(" Then
        argString = Mid(argString, 2)
    End If
    
    argString = Trim(argString)
    If Len(argString) = 0 Then Exit Sub
    If Left(argString, 1) = ")" Then Exit Sub
    
    Dim ca As CArg
    
    If InStr(argString, ",") < 1 Then
        Set ca = New CArg
        ca.ParseArg argString
        args.Add ca
    Else
        tmp = Split(argString, ",")
        For Each X In tmp
            Set ca = New CArg
            ca.ParseArg X
            args.Add ca
        Next
    End If
    
End Sub



Function AddLine(X)
    push m_body, X
End Function

'finalizes the loading of the procedure..
Function isEndProc(ByVal X, owner As CRefactor) As Boolean
    
    ft = ft_notset
    X = LCase(X)
    X = Replace(X, vbTab, Empty)
    
    While InStr(X, "  ") > 0
        X = Replace(X, "  ", " ")
    Wend
    
    X = Trim(X)
    If X = "end sub" Then
        ft = ft_sub
        isEndProc = True
        LastLine = "End Sub"
    End If
    
    If X = "end function" Then
        ft = ft_function
        isEndProc = True
        LastLine = "End Function"
    End If
    
    If X = "end property" Then
        ft = ft_property
        isEndProc = True
        LastLine = "End Property"
    End If
    
    'now we scan it to catalog local vars
    If isEndProc And Not AryIsEmpty(m_body) Then
        For i = 0 To UBound(m_body)
            loadDims m_body(i), localVars
            loadNonExplicitVars m_body(i), localVars, Me.name, Me.args, owner.gVars
        Next
    End If
    
    
End Function

Sub RebuildProto()
    
    Dim ca As CArg
    Dim t As CToken
    Dim c As Collection
    
    n_body() = m_body() 'make a copy of the body text for mod next (this func only called once..)
    Set c = TokenizeLine(rawProto)
    For Each t In c
        If LCase(t.token) = LCase(name) Then
            t.token = newName
        End If
        For Each ca In args
            If LCase(t.token) = LCase(ca.name) Then
                t.token = ca.newName
            End If
        Next
    Next
    
    Me.newProto = ReassembleTokenizedString(c)
    
End Sub

Sub RenameTokens(oldName, newName)
    
    'this func called many times..
    Dim t As CToken
    Dim c As Collection
    Dim i As Long
    Dim changed As Boolean
    
    If AryIsEmpty(n_body) Then Exit Sub
    
    For i = 0 To UBound(n_body)
        changed = False
        Set c = TokenizeLine(n_body(i))
        For Each t In c
            If LCase(t.token) = LCase(oldName) Then
                t.token = newName
                changed = True
            End If
        Next
        If changed Then
            n_body(i) = ReassembleTokenizedString(c)
        End If
    Next
    
End Sub

