VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRefactor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public gVars As New Collection   'of CArg
Public methods As New Collection 'of CFunc - this one is easy..
Private m_GlobalScript() As String

'parsing out variables comes in several flavors..
'dim'ed, un declared, global, local to a function, and as function arguments..
'so to track variables..you need to parse out the functions first and record attributes..

Property Get GlobalScript() As String
    GlobalScript = Join(m_GlobalScript, vbCrLf)
End Property

'this assumes the code has already been formatted to one command per line
Function Analyze(code As String)
    
    Dim inProc As Boolean
    Dim cf As CFunc
    Dim i As Long
    Dim x As String
    Dim tmp() As String
    
    tmp = Split(code, vbCrLf)
    
    Erase m_GlobalScript
    Set gVars = New Collection
    Set methods = New Collection

    For i = 0 To UBound(tmp)
        
        x = tmp(i)
        'If InStr(1, x, "StartCalendar") > 0 Then Stop
        'If InStr(x, vbCr) > 0 Then Stop
        'If InStr(x, vbLf) > 0 Then Stop
        
        If inProc Then
            If cf.isEndProc(x, Me) Then
                inProc = False
                methods.Add cf
                Set cf = Nothing
                GoTo nextLine
            Else
                cf.AddLine x
                GoTo nextLine
            End If
        Else
            loadDims x, gVars
            loadNonExplicitVars x, gVars
        End If
        
        If isFuncStart(x) Then
           inProc = True
           Set cf = New CFunc
           cf.SetFirstLine x, i
           push m_GlobalScript, cf.gsMarker
        Else
            If Not inProc Then
                push m_GlobalScript, x
            End If
        End If
        
nextLine:
         
    Next
    
End Function

Sub RenameAll()
    
    Dim ca As CArg, cf2 As CFunc, cf As CFunc
    Dim i As Long, j As Long

    'generate new names for all methods and thier arguments..
    For Each cf In Me.methods
        cf.newName = "func_" & Format(i, "000")
        For Each ca In cf.args
            ca.newName = "arg" & Format(j, "00")
            j = j + 1
        Next
        j = 0
        For Each ca In cf.localVars
            ca.newName = "var" & Format(j, "00")
            j = j + 1
        Next
        i = i + 1
    Next
    
    i = 0
    For Each ca In Me.gVars
       ca.newName = "gvar_" & Format(i, "00")
       i = i + 1
    Next
    
End Sub

Function RebuildGlobalScript() As String
    
    Dim cf As CFunc
    Dim tmp() As String

    Dim t As CToken
    Dim c As Collection
    Dim i As Long
    Dim changed As Boolean
    Dim ca As CArg
    
    For i = 0 To UBound(m_GlobalScript)
        changed = False
        Set c = TokenizeLine(m_GlobalScript(i))
        If c.Count > 0 Then
            For Each ca In gVars
                For Each t In c
                    If LCase(t.token) = LCase(ca.name) Then
                        t.token = ca.newName
                        changed = True
                    End If
                Next
            Next
            For Each cf In methods
                For Each t In c
                    If LCase(t.token) = LCase(cf.name) Then
                        t.token = cf.newName
                        changed = True
                    End If
                Next
            Next
        End If
        If changed Then
            m_GlobalScript(i) = ReassembleTokenizedString(c)
        End If
    Next
    
    tmp = m_GlobalScript
    For i = 0 To UBound(tmp)
        For Each cf In Me.methods
            If tmp(i) = cf.gsMarker Then
                tmp(i) = cf.newBody
                Exit For
            End If
        Next
    Next
    
    RebuildGlobalScript = Join(tmp, vbCrLf)
    
End Function
