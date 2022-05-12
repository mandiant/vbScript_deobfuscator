Attribute VB_Name = "modParse"
Global dlg As New clsCmnDlg2

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Const EM_LINESCROLL = &HB6
Const EM_GETFIRSTVISIBLELINE = &HCE
Const MAX_LONG As Long = 2147483647
Const EM_LINELENGTH = &HC1

Function isFuncStart(ByVal x) As Boolean

    'must pass: function StartCalENDer()
    
    x = Trim(LCase(x))
    a = InStr(1, x, "function")
    b = InStr(1, x, "sub")
    c = InStr(1, x, "property")
    
    Dim low  As Long
    low = MAX_LONG
    If a > 0 Then low = a
    If b > 0 And b < low Then low = b
    If c > 0 And c < low Then low = c
    
    If low = MAX_LONG Then Exit Function
    
    'todo: parse public, private, friend
    
    If low > 0 Then
           d = InStr(1, x, "end", vbTextCompare)
           e = InStr(1, x, "exit", vbTextCompare)
           If d > 0 And d < low Then Exit Function
           If e > 0 And e < low Then Exit Function
    End If
    
    isFuncStart = True

End Function

'assumes we are one statement per line already..
'should we tokenize? this could be delicate..
'I should study and rip the tokenizer from something like freebasic...
Sub loadDims(ByVal x, cVars As Collection)
    Dim marker As String
    
    If Len(x) = 0 Then Exit Sub
    If InStr(1, x, "function ", vbTextCompare) > 0 Then Exit Sub
    If InStr(1, x, "sub ", vbTextCompare) > 0 Then Exit Sub
    If InStr(1, x, "property ", vbTextCompare) > 0 Then Exit Sub
    
    marker = "dim "
    If InStr(1, x, "const ", vbTextCompare) > 0 Then marker = Empty
    
    x = Replace(x, "public ", marker, , , vbTextCompare)
    x = Replace(x, "private ", marker, , , vbTextCompare)
    x = Trim(x)
    
    If Len(marker) <> 0 Then
        internal_loadDims x, cVars, "dim "
    Else
        internal_loadDims x, cVars, "const "
    End If
            
End Sub

 Sub internal_loadDims(ByVal x, cVars As Collection, Optional ByVal marker = "dim ")
    
    Dim ca As CArg
    
    a = InStr(1, x, marker, vbTextCompare)
    
    If a > 1 Then
        prev = Mid(x, a - 1, 1)
        If prev <> " " And prev <> vbTab Then Exit Sub
    End If
    
    If a > 0 Then
        x = Mid(x, a + Len(marker))
        b = InStr(x, ",") 'bug: const x = "this,that"
        If b > 0 Then
            tmp = Split(x, ",")
            For Each Y In tmp
                Y = Trim(Y)
                If Len(Y) > 0 Then
                    Set ca = New CArg
                    ca.ParseArg Y
                    cVars.Add ca
                End If
            Next
        Else
            Set ca = New CArg
            ca.ParseArg Trim(x)
            cVars.Add ca
        End If
    End If
            
            
End Sub

'assumes we are one statement per line already..
Sub loadNonExplicitVars(ByVal x, cVars As Collection, Optional ignoreName As String, Optional ignoreArgs As Collection, Optional ignoreGlobals As Collection)
    
    Dim ca As CArg
    
    If InStr(1, x, "const ", vbTextCompare) > 0 Then Exit Sub
    
    x = StripQuotedStringAndCmt(x)
    a = InStr(1, x, "=", vbTextCompare)
    If a < 2 Then Exit Sub
    
    tmp = Trim(Mid(x, 1, a - 1))
    b = InStrRev(tmp, " ")
    If b > 0 Then
        tmp = Trim(Mid(tmp, b))
    End If
    
    If Len(tmp) > 0 Then
        
        'when parsing function localVars, ignore function name itself if used for return val
        If Len(ignoreName) > 0 Then
            If LCase(tmp) = LCase(ignoreName) Then Exit Sub
        End If
        
        If Not ignoreArgs Is Nothing Then
            For Each ca In ignoreArgs
               If LCase(ca.name) = LCase(tmp) Then Exit Sub
               If LCase(ca.newName) = LCase(tmp) Then Exit Sub
            Next
        End If
        
        If Not ignoreGlobals Is Nothing Then
            For Each ca In ignoreGlobals
               If LCase(ca.name) = LCase(tmp) Then Exit Sub
               If LCase(ca.newName) = LCase(tmp) Then Exit Sub
            Next
        End If
        
        If Not isValidVarExtraction(tmp) Then Exit Sub
        
        If Not VarExists(tmp, cVars) Then
            Set ca = New CArg
            ca.name = tmp
            cVars.Add ca
        End If
        
    End If
    
End Sub

Function isValidVarExtraction(name) As Boolean
    
    If Not CTokenizer.isInit Then CTokenizer.Init
    
    If AnyofTheseInstr(name, Join(seps, ",")) Then
        Exit Function
    End If
    
    isValidVarExtraction = True
    
End Function

Function VarExists(name, cVars As Collection) As Boolean
    Dim ca As CArg
    For Each ca In cVars
        If LCase(name) = LCase(ca.name) Then
            VarExists = True
        End If
    Next
End Function

Function StripQuotedStringAndCmt(sin) As String
    
    tmp = Trim(Replace(sin, vbTab, Empty))
    
    If Len(tmp) = 0 Then Exit Function
    If Left(tmp, 1) = "'" Then Exit Function
    
    tmp = Replace(tmp, """""", Chr(4)) 'replace double quotes as chr(4) (embedded double quotes and empty strings)
    tmp = Split(tmp, """")
    
    Dim ret() As String
    
    For i = 0 To UBound(tmp)
        If Not inString Then
            push ret, tmp(i)
        End If
        inString = Not inString
    Next

    tmp = Join(ret, "")
    tmp = Replace(tmp, Chr(4), "")
    
    'not in a quoted string, so must be a comment.
    a = InStr(1, tmp, "'")
    If a > 2 Then
        tmp = Mid(tmp, 1, a - 1)
    End If
    
    StripQuotedStringAndCmt = tmp
    
End Function



Function AryIsEmpty(ary) As Boolean
  On Error GoTo oops
    i = UBound(ary)  '<- throws error if not initalized
    AryIsEmpty = False
  Exit Function
oops: AryIsEmpty = True
End Function


Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo Init
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
Init:     ReDim ary(0): ary(0) = value
End Sub
 
 
Sub GotoLine(t As Object, ByVal lNo As Integer, Optional ByRef lLen As Long)
    
    Dim charIndex As Long
    Const EM_LINEINDEX = &HBB
    
    lNo = lNo '- 1
    charIndex = SendMessage(t.hwnd, EM_LINEINDEX, ByVal lNo, ByVal CLng(0))
    If charIndex = -1 Then Exit Sub
    
    lLen = SendMessage(t.hwnd, EM_LINELENGTH, charIndex, 0)
    
    t.selLength = 0
    t.SetFocus
    t.selStart = charIndex

End Sub

Function AnyofTheseInstr(sin, sCmp) As Boolean
    Dim tmp() As String, i As Integer
    tmp() = Split(sCmp, ",")
    For i = 0 To UBound(tmp)
        tmp(i) = Trim(tmp(i))
        If Len(tmp(i)) > 0 And InStr(1, sin, tmp(i), vbTextCompare) > 0 Then
            AnyofTheseInstr = True
            Exit Function
        End If
    Next
End Function

Function AnyofTheseInstr2(data, match, Optional compare As VbCompareMethod = vbTextCompare, Optional divider = ",") As Boolean
    Dim tmp() As String
    Dim x
    Dim b() As Byte, i As Long
    
    If Len(divider) > 0 Then
        tmp = Split(match, divider)
    Else
        b() = StrConv(match, vbFromUnicode, LANG_US)
        For i = 0 To UBound(b)
            push tmp, Chr(b(i))
        Next
    End If
    
    For Each x In tmp
        If InStr(1, data, x, compare) > 0 Then
            AnyofTheseInstr2 = True
            Exit Function
        End If
    Next
    
End Function

Function IsIde() As Boolean
' Brad Martinez  http://www.mvps.org/ccrp
    On Error GoTo out
    Debug.Print 1 / 0
out: IsIde = Err
End Function

Function ReadFile(filename)
  f = FreeFile
  temp = ""
   Open filename For Binary As #f        ' Open file.(can be text or image)
     temp = Input(FileLen(filename), #f) ' Get entire Files data
   Close #f
   ReadFile = temp
End Function



Function FileExists(path As String) As Boolean
  On Error GoTo hell
    
  If Len(path) = 0 Then Exit Function
  If Right(path, 1) = "\" Then Exit Function
  If Dir(path, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then FileExists = True
  
  Exit Function
hell: FileExists = False
End Function



Function GetParentFolder(path) As String
    tmp = Split(path, "\")
    ub = tmp(UBound(tmp))
    GetParentFolder = Replace(Join(tmp, "\"), "\" & ub, "")
End Function



Sub WriteFile(path, it)
    f = FreeFile
    Open path For Output As #f
    Print #f, it
    Close f
End Sub

