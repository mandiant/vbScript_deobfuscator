Attribute VB_Name = "modSyntaxHighlighting"
Option Explicit

'Copyright David Zimmer <dzzie@yahoo.com> - Oct 19 03

Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Private Declare Function SendMessageStr Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long


Private Type CHARFORMAT2
    cbSize As Integer
    wPad1 As Integer
    dwMask As charFormatMasks 'Long
    dwEffects As charFormatEffects 'Long
    yHeight As Long
    yOffset As Long
    crTextColor As Long
    bCharSet As Byte
    bPitchAndFamily As Byte
    szFaceName(0 To 32 - 1) As Byte
    wPad2 As Integer
    'Additional stuff supported by RICHEDIT20
    wWeight As Integer     'Font weight (LOGFONT value)
    sSpacing As Integer    'Amount to space between letters
    crBackColor As Long    'Background color
    lLCID As Long          'Locale ID
    dwReserved As Long     'Reserved. Must be 0
    sStyle As Integer      'Style handle
    wKerning As Integer    'Twip size above which to kern char pair
    bUnderlineType As Byte 'Underline type
    bAnimation As Byte     'Animated text like marching ants
    bRevAuthor As Byte     'Revision author index
    bReserved1 As Byte
End Type

Private Const WM_USER = &H400

Private Enum tmMsgs
        EM_CHARFROMPOS = &HD7
        EM_UNDO = &HC7
        EM_CANUNDO = &HC6
        EM_SETWORDBREAKPROC = &HD0
        EM_SETTABSTOPS = &HCB
        EM_SETSEL = &HB1
        EM_SETRECTNP = &HB4
        EM_SETRECT = &HB3
        EM_SETREADONLY = &HCF
        EM_SETPASSWORDCHAR = &HCC
        EM_SETMODIFY = &HB9
        EM_SCROLLCARET = &HB7
        EM_SETHANDLE = &HBC
        EM_SETCHARFORMAT = (WM_USER + 68)
        EM_SCROLL = &HB5
        EM_REPLACESEL = &HC2
        EM_LINESCROLL = &HB6
        EM_LINELENGTH = &HC1
        EM_LINEINDEX = &HBB
        EM_LINEFROMCHAR = &HC9
        EM_LIMITTEXT = &HC5
        EM_GETWORDBREAKPROC = &HD1
        EM_GETTHUMB = &HBE
        EM_GETRECT = &HB2
        EM_GETSEL = &HB0
        EM_GETPASSWORDCHAR = &HD2
        EM_GETMODIFY = &HB8
        EM_GETLINECOUNT = &HBA
        EM_GETLINE = &HC4
        EM_GETHANDLE = &HBD
        EM_GETFIRSTVISIBLELINE = &HCE
        EM_FMTLINES = &HC8
        EM_EMPTYUNDOBUFFER = &HCD
        EM_SETMARGINS = &HD3
End Enum

Private Enum charFormatMasks
    CFM_BOLD = &H1&
    CFM_ITALIC = &H2&
    CFM_UNDERLINE = &H4&
    CFM_STRIKEOUT = &H8&
    CFM_PROTECTED = &H10&
    CFM_LINK = &H20&
    CFM_SIZE = &H80000000
    CFM_COLOR = &H40000000
    CFM_FACE = &H20000000
    CFM_OFFSET = &H10000000
    CFM_CHARSET = &H8000000
    CFM_BACKCOLOR = &H4000000
End Enum

Private Enum charFormatEffects
    CFE_BOLD = &H1&
    CFE_ITALIC = &H2&
    CFE_UNDERLINE = &H4&
    CFE_STRIKEOUT = &H8&
    CFE_PROTECTED = &H10&
    CFE_LINK = &H20&
    CFE_AUTOCOLOR = &H40000000
End Enum

Private LockCount As Integer
Private Const SCF_SELECTION = &H1

Private mrtf As RichTextBox

Private vbTokens() As String
Private jsTokens() As String
Private lvbTokens() As String
Private ljsTokens() As String
Private bvbtokens() As Byte
Private bjsTokens() As Byte

Private tokensInitalized As Boolean

Private Sub InitalizeTokens()
    
    tokensInitalized = True
    
    vbTokens() = Split("If,Then,Else,ElseIf,Case,Default,With,End," & _
                       "Select,cStr,Not,Exit,Function,Sub,And,Or,Xor," & _
                       "For,While,Next,Wend,Do,Loop,Until,Mid,InStr," & _
                       "Left,Trim,LTrim,Right,UBound,LBound,Len,Split," & _
                       "Call,True,False,Set,To,Each,In,Is,Nothing,Dim," & _
                       "Redim,Preserve,On,Error,Resume,True,False,IsArray," & _
                       "IsObject,IsNumeric,Const,String,Integer,Long,Byte," & _
                       "Variant,Double,Single,Currency,As,Property,Get,Let," & _
                       "Set,Boolean,ByVal,ByRef,Public,Declare,Private,option,explicit", ",")
                    
    jsTokens() = Split("if,else,switch,new,var,function,eval,break,exit," & _
                       "for,while,case,default,true,false,NaN", ",")
    
    ReDim lvbTokens(UBound(vbTokens))
    ReDim ljsTokens(UBound(jsTokens))
    ReDim bvbtokens(UBound(vbTokens))
    ReDim bjsTokens(UBound(jsTokens))

   Dim i As Integer

    For i = 0 To UBound(vbTokens)
        lvbTokens(i) = LCase(vbTokens(i))
        bvbtokens(i) = LCaseLeft1(vbTokens(i))
    Next

    For i = 0 To UBound(jsTokens)
        ljsTokens(i) = LCase(jsTokens(i))
        bjsTokens(i) = LCaseLeft1(jsTokens(i))
    Next
    
End Sub


Sub SyntaxHighlight(rtf As RichTextBox, Optional isVbs As Boolean = True, Optional lineIndex As Long = -1)
    Dim i As Long, topLine As Long
    
    
    On Error GoTo hell

    Set mrtf = rtf
    topLine = TopLineIndex
    LockUpdate
        
    If lineIndex < 0 Then
        'clear all comment formatting in one bulk operation
        rtf.selStart = 1
        rtf.selLength = Len(rtf.Text)
        rtf.SelColor = vbBlack
        rtf.selLength = 0

        For i = 0 To lineCount
            SyntaxHighlightLine i, isVbs, False
        Next
        rtf.selStart = 1
        rtf.SelColor = vbBlack
    Else
        SyntaxHighlightLine lineIndex, isVbs
        'rtf.selStart = .IndexOfFirstCharOnLine(lineIndex + 1)
        'rtf.SelColor = vbBlack
    End If
   
hell:
    ScrollToLine topLine
    'mrtf.selLength = 0
    UnlockUpdate
    
End Sub

 


Sub SyntaxHighlightLine(lineIndex As Long, isVbs As Boolean, Optional MakeSureClear As Boolean = True)
    
    'Exit Sub
    
    Dim tmp As String, i As Integer, j As Integer
    Dim commentStart As Integer, commentlength As Integer
    Dim lineStart As Long
    Dim words() As String
    Dim tokens() As String
    Dim lTokens() As String
    Dim bTokens() As Byte
    
    If Not tokensInitalized Then InitalizeTokens
    
    If isVbs Then
        tokens() = vbTokens()
        lTokens() = lvbTokens()
        bTokens() = bvbtokens()
    Else
        tokens() = jsTokens()
        lTokens() = ljsTokens()
        bTokens() = bjsTokens()
    End If
    
   
        tmp = GetLine(lineIndex)
        If Len(Trim(tmp)) = 2 Then Exit Sub 'account for the vbcrlf
        
        lineStart = IndexOfFirstCharOnLine(lineIndex)
        
        If MakeSureClear Then  'clear previous formatting
            PreformHighlight lineStart, Len(tmp), vbBlack
        End If
        
        commentStart = CommentStartChar(tmp, isVbs)
        commentlength = Len(tmp) - commentStart
        
        If commentStart > 0 Then
            tmp = Mid(tmp, 1, commentStart)
            PreformHighlight (lineStart + commentStart), commentlength, RGB(0, &H88, 0)
        End If
        
        tmp = Replace(tmp, vbCrLf, Empty) 'cleanup
        tmp = Replace(tmp, vbTab, " ") 'cleanup
        tmp = Replace(tmp, "(", " ") 'valid word divider
        tmp = Replace(tmp, ",", " ") 'valid word divider
        tmp = Replace(tmp, "{", " ") 'valid word divider
        tmp = Replace(tmp, ";", " ") 'valid word divider
        tmp = Replace(tmp, ":", " ") 'valid word divider
        
        
        'tmp is not whole line no comments
        If Len(Trim(tmp)) = 0 Then Exit Sub
        
        'now block out all the quoted strings
        'word & character indexes remain unchanged
        RemoveQuotedStrings tmp
        
        words() = Split(tmp, " ")
        
        Dim wordStartIndex As Integer
        
        wordStartIndex = 0
        Dim lCasethisWord As String
        Dim thisWordFirstLetter As Byte
        For i = 0 To UBound(words)
            If Len(words(i)) > 0 Then
                lCasethisWord = LCase(words(i))
                thisWordFirstLetter = LCaseLeft1(words(i))
                'Debug.Print lCasethisWord & ":" & thisWordFirstLetter
                For j = 0 To UBound(tokens)
                    'Debug.Print LCase(tokens(j)) & " " & LCase(words(i))
                     If bTokens(j) = thisWordFirstLetter Then
                        If lTokens(j) = lCasethisWord Then
                            PreformHighlight lineStart + wordStartIndex, Len(words(i)), RGB(0, 0, &H88), tokens(j)
                        End If
                     End If
                Next
                wordStartIndex = wordStartIndex + Len(words(i))
            End If
            wordStartIndex = wordStartIndex + 1 'for the spaces
        Next
                   
   
End Sub

Private Function LCaseLeft1(s As String) As Byte
    If Len(s) = 0 Then Exit Function
    LCaseLeft1 = AscW(s)
    If LCaseLeft1 < 97 Then
        LCaseLeft1 = LCaseLeft1 + (97 - 65)
    End If
End Function

Private Sub PreformHighlight(selStart, selLength, color, Optional selText = "")
 
reSelect:
        mrtf.selStart = IIf(selStart >= 0, selStart, 1)
        mrtf.selLength = IIf(selLength > 0, selLength, 1)
        
        If Len(selText) > 0 Then
            mrtf.selText = selText 'for proper case
            selText = ""
            GoTo reSelect
        End If
        
        Dim l As Long
        l = Len(mrtf.Text)
        If mrtf.selStart + mrtf.selLength >= l Then Exit Sub
        
        'Open "C:\log.txt" For Append As 1
        'Print 1, mrtf.selStart & " " & mrtf.selLength
        'Close 1
        
        HighLightSelection vbWhite, color
 
    
End Sub

Private Sub HighLightSelection(Optional bgColor = vbYellow, Optional fgColor = vbBlack)
    Dim udtCharFormat As CHARFORMAT2
    With udtCharFormat
        .cbSize = Len(udtCharFormat)
        .dwMask = CFM_BACKCOLOR Or CFM_COLOR
        .crBackColor = bgColor
        .crTextColor = fgColor
    End With
    SendMessage mrtf.hwnd, EM_SETCHARFORMAT, SCF_SELECTION, udtCharFormat
End Sub

Private Sub RemoveQuotedStrings(sin As String)
    Dim dq As Integer
    Dim match As Integer, sLen As Integer
    Dim tmp As String
    
    dq = InStr(sin, """")
    While dq > 0
        dq = InStr(sin, """")
        If dq < 1 Then Exit Sub
        match = InStr(dq + 1, sin, """") 'find its closing dq
        If match < 1 Then Exit Sub 'err.raise turn line red?
        sLen = match - dq + 1  'entire length of the quoted string
        tmp = Mid(sin, 1, dq - 1) & String(sLen, "-") & Mid(sin, match + 1, Len(sin))
        sin = tmp
    Wend
    
End Sub

Private Function CommentStartChar(ByVal sLine As String, Optional isVbs As Boolean) As Integer
    Dim commentChar As String
    Dim sq As Integer, dq As Integer, startAt As Integer
    
    commentChar = IIf(isVbs, "'", "//")
    
    If Not isVbs Then
        'so we dont have to deal with the possibility of single or double quoted strings
        sLine = Replace(sLine, "'", " ")
    End If
    
    startAt = 1
    
top:
    sq = InStr(startAt, sLine, commentChar)
    dq = InStr(startAt, sLine, """")
    
    If sq < 1 Then Exit Function
    
    If dq > 1 And dq < sq Then
        'we are in a quoted string, find the end quote and change startAt
        dq = InStr(dq + 1, sLine, """")
        If dq < 1 Then 'no close quote found? exit
            Exit Function
        Else
            startAt = dq + 1
            GoTo top
        End If
    End If
    
    'by the time we get here, sq should be at the first char of comment block
    CommentStartChar = sq - IIf(sq = 1, 0, 1)
    'think this leaves a bug but acceptable for now
End Function


Public Function LMouseDown() As Boolean
    GetAsyncKeyState vbKeyLButton
    LMouseDown = Not (GetAsyncKeyState(vbKeyLButton) And &HFFFF) = 0
End Function

Public Function RMouseDown() As Boolean
    GetAsyncKeyState vbKeyRButton
    RMouseDown = Not (GetAsyncKeyState(vbKeyRButton) And &HFFFF) = 0
End Function

Private Function GetLine(lineNo As Long) As String
    'When retrieving a line into a buffer the first word
    'of the buffer specifies the max number of bytes to read,
    'so one has to guess the maximum line length.  When the bytes
    'are read into the buffer the first word gets overwritten.
    'Remember that lines in a TextBox are numbered starting at zero.

    Dim lret As Long
    Dim strBuffer As String
    Dim intLineLen As Integer
    Dim lngLength As Long
    Dim lFirstCharIndexForLine As Long
    
    lFirstCharIndexForLine = SendMessage(mrtf.hwnd, EM_LINEINDEX, lineNo, 0)
    If lFirstCharIndexForLine < 1 Then Exit Function

    lngLength = SendMessage(mrtf.hwnd, EM_LINELENGTH, lFirstCharIndexForLine, 0)
    If lngLength < 1 Then Exit Function
    
    strBuffer = Space(lngLength + 20) 'max width possible for line
    intLineLen = SendMessageStr(mrtf.hwnd, EM_GETLINE, lineNo, strBuffer)
    GetLine = Left(strBuffer, intLineLen)


End Function

Private Function IndexOfFirstCharOnLine(lNo As Long)
    IndexOfFirstCharOnLine = SendMsg(EM_LINEINDEX, lNo)
End Function

Private Sub LockUpdate()
    If LockCount = 0 Then LockWindowUpdate mrtf.hwnd
    LockCount = LockCount + 1
End Sub

Private Sub UnlockUpdate()
    LockCount = LockCount - 1
    If LockCount = 0 Then
        LockWindowUpdate 0
    End If
End Sub

Private Function SendMsg(Msg As tmMsgs, Optional wParam As Long = 0, Optional lParam = 0) As Long
    SendMsg = SendMessage(mrtf.hwnd, Msg, wParam, lParam)
End Function

Private Function lineCount() As Long
    lineCount = SendMessage(mrtf.hwnd, EM_GETLINECOUNT, 0, ByVal 0&)
End Function

Private Function TopLineIndex() As Long
    TopLineIndex = SendMessage(mrtf.hwnd, EM_GETFIRSTVISIBLELINE, 0, ByVal 0&)
End Function

Private Sub ScrollToLine(x As Long)
     x = x - TopLineIndex
     ScrollIncremental , x
End Sub

Private Sub ScrollIncremental(Optional horz As Long = 0, Optional vert As Long = 0)
    'lParam&  The low-order 2 bytes specify the number of vertical
    '          lines to scroll. The high-order 2 bytes specify the
    '          number of horizontal columns to scroll. A positive
    '          value for lParam& causes text to scroll upward or to the
    '          left. A negative value causes text to scroll downward or
    '          to the right.
    ' r&       Indicates the number of lines actually scrolled.
    
    Dim r As Long
    r = CLng(&H10000 * horz) + vert
    r = SendMessage(mrtf.hwnd, EM_LINESCROLL, 0, ByVal r)

End Sub

