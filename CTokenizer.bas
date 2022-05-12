Attribute VB_Name = "CTokenizer"
Private isSeparator(0 To 255) As Boolean
Public isInit As Boolean
Public seps() As String

Global Const funcOpenMarker As String = "[--------"


Private startTime As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long

Sub StartBenchMark()
    startTime = GetTickCount()
End Sub

Function EndBenchMark() As String
    Dim endTime As Long, loadTime As Long
    endTime = GetTickCount()
    loadTime = endTime - startTime
    EndBenchMark = loadTime / 1000 & " seconds"
End Function

Public Sub Init()

   Const sSeparators = vbTab & " ,.:;!?""()=-><+&" & vbCrLf
   Dim nI As Integer
   
   For nI = 1 To Len(sSeparators)
      isSeparator(Asc(Mid$(sSeparators, nI, 1))) = True
      push seps, Mid$(sSeparators, nI, 1)
   Next

   isInit = True

End Sub

Function ReassembleTokenizedString(c As Collection) As String
    
    Dim t As CToken
    Dim ret() As String
    
    For Each t In c
        push ret, t.token & t.seperator
    Next
    
    ReassembleTokenizedString = Join(ret, "")
        
    'If InStr(ReassembleTokenizedString, "arg00 = arg00&") > 0 Then Stop
    
End Function

Function TokenizeLine(sin) As Collection
    
    Dim inString As Boolean
    Dim v As New Collection
    Dim v2 As Collection
    Dim t As CToken
    
    Set TokenizeLine = v
    
    If Left(sin, Len(funcOpenMarker)) = funcOpenMarker Then Exit Function
    
    tmp = Replace(sin, """""", Chr(4)) 'replace double quotes as chr(4) (embedded double quotes and empty strings)
    tmp = Split(tmp, """")
    
    For i = 0 To UBound(tmp)
        If Not inString Then
            'Debug.Print "noInstr: " & tmp(i)
            'Debug.Print "Tokens for line: " & sin
            Set v2 = internal_TokenizeLine(Replace(tmp(i), Chr(4), """"""))
            For Each vv In v2
                v.Add vv
            Next
        Else
            'Debug.Print "Instr: " & tmp(i)
            Set t = New CToken
            'If InStr(tmp(i), Chr(4)) Then Stop
            t.AddString Replace(tmp(i), Chr(4), """""")
            v.Add t
        End If
        inString = Not inString
    Next

End Function

Function internal_TokenizeLine(x) As Collection
    
    Dim v As New Collection
    Dim t As CToken
    Dim pos As Long
    
    If Not isInit Then Init
    
    Do While 1
        Set t = NextToken(x, pos)
        If t Is Nothing Then Exit Do
        pos = pos + Len(t.token) + Len(t.seperator)
        'Debug.Print t.token & " " & t.seperator
        v.Add t
        DoEvents
    Loop

    If pos < Len(x) Then
        Set t = New CToken
        t.token = Mid(x, pos)
        v.Add t
    End If
    
    Set internal_TokenizeLine = v
    
End Function

Private Function NextToken(x, ByRef startAt) As CToken
    
    Dim i As Long
    Dim nt As CToken
    Dim b As Byte
    
    If startAt = 0 Then startAt = 1
    If startAt > Len(x) Then Exit Function
    
    For i = startAt To Len(x)
        b = Asc(Mid(x, i, 1))
        If isSeparator(b) Then Exit For
    Next
            
    Set nt = New CToken
    
    If i = Len(x) + 1 Then         'read to end of line with no more tokens..
        nt.token = Mid(x, startAt)
        nt.seperator = Empty
    Else
        nt.token = Mid(x, startAt, i - startAt)
        nt.seperator = Chr(b)
    End If
    
    Set NextToken = nt
    
End Function

