VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CArg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public name As String
Public argType As String
Public newName As String
Public isArray As Boolean
Public isOptional As Boolean
Public isByVal As Boolean
Public defValue
Public isConst As Boolean

'todo: test optional byref byval and add support for defaultValue

Sub ParseArg(ByVal arg)
    
    Dim i As Long, X
    On Error Resume Next
    
    arg = Replace(arg, vbTab, Empty)
    arg = Trim(arg)
    
    If InStr(arg, " ") < 1 Then
        name = arg
        Exit Sub
    End If
    
    While InStr(arg, "  ") > 0
        arg = Replace(arg, "  ", " ")
    Wend
    
    tmp = Split(arg, " ")
    
    name = tmp(i)
    
    If LCase(name) = "optional" Then
        isOptional = True
        i = i + 1
        name = tmp(i)
    End If
    
    If LCase(name) = "byval" Then
        isByVal = True
        i = i + 1
        name = tmp(i)
    ElseIf LCase(name) = "byref" Then
        i = i + 1
        name = tmp(i)
    End If
    
    
    For i = i To UBound(tmp)
        X = tmp(i)
        If LCase(X) = "as" Then
            If i <= UBound(tmp) - 1 Then argType = tmp(i + 1)
        End If
        If X = "=" Then
            If i <= UBound(tmp) - 1 Then defValue = tmp(i + 1)
        End If
    Next
        
    a = InStr(name, "(")
    If a > 2 Then
        name = Mid(name, 1, a - 1)
        isArray = True
    End If
    
End Sub


