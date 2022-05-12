VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CIndenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' #VBIDEUtils#************************************************************
' * Programmer Name  : removed
' * Web Site         : http://www.ppreview.net
' * E-Mail           : removed
' * Date             : 25/11/97
' * Time             : 11:17
' * Module Name      : class_Indenter
' * Module Filename  : Indent.cls
' **********************************************************************
' * Comments         : Indent code in an efficient way
' *
' *
' **********************************************************************

'Option Explicit
Public Event Progress(pcent As Integer)

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, Source As Any, ByVal bytes As Long)
Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (dest As Any, ByVal bytes As Long)

Private Const STARTBLOCK = 0
Private Const UNALTERED = 1
Private Const MIDBLOCK = 2
Private Const ENDBLOCK = 3
Private Const NOINDENTBLOCK = 4
Private Const CASEBLOCK = 5
Private Const ENDCASEBLOCK = 6
Private Const NOALIGNCOMMENTLOCK = 7
Private Const UNDERSCORECHAR = 8
Private Const LABELLINE = 9
Private Const BLANKLINE = 10
Private Const VARIABLELINE = 11

Private Type LINEINFOS
   sTextLine            As String ' not used
   nNumTabs             As Integer
End Type

Private isSeparator(0 To 255) As Boolean
Private sAllLines()     As String
Private byteArrayString() As Byte

Private nIndentSpaces   As Integer
Private bIndentVariable As Boolean
Private bIndentCase     As Boolean
Private bUseTabs        As Boolean
Private bIndentProc     As Boolean
Private bIndentCmt      As Boolean
Private sCharIndent     As String * 1

Private sInProc         As String
Private sInCode         As String
Private sMidCode        As String
Private sOutProc        As String
Private sOutCode        As String

Private sBuffer()       As LINEINFOS

Private gbUnderscore    As Boolean
Private gsFirstWord     As String
Private gsSecondWord    As String


Private Sub Class_Initialize()

   Const sSeparators = vbTab & " ,.:;!?""()=-><+&" & vbCrLf
   Dim nI As Integer
   
   For nI = 1 To Len(sSeparators)
      isSeparator(Asc(Mid$(sSeparators, nI, 1))) = True
   Next

   SetOptions

End Sub

Sub SetOptions(Optional useTabs As Boolean = True, _
                Optional IndentProc As Boolean = True, _
                Optional IndentCmt As Boolean = True, _
                Optional IndentCase As Boolean = True, _
                Optional IndentSpaces As Integer = 3, _
                Optional IndentVariable As Boolean = True _
)
                
   bUseTabs = useTabs
   bIndentProc = IndentProc
   bIndentCmt = IndentCmt
   bIndentCase = IndentCase
   nIndentSpaces = IndentSpaces
   bIndentVariable = IndentVariable

   If bUseTabs Then
      sCharIndent = vbTab
   Else
      sCharIndent = Chr$(32)
   End If

   If bIndentCase Then
      sInProc = "Select Case*"
      sOutCode = "End Select*"
   Else
      sInProc = Empty
      sOutCode = Empty
   End If
              
End Sub

Public Function IndentVBCode(sVBCode As String, _
                             sAllLines() As String, _
                             Optional bStripContinuations As Boolean = True, _
                             Optional bExpandColonCmds As Boolean = True, _
                             Optional bStandardize As Boolean = True, _
                             Optional bAddNewLines As Boolean = True _
                ) As String
                
                
   Dim sin As String
   Dim addit As Boolean
   Dim j As Long
   
   If Len(sVBCode) = 0 Then Exit Function

   sin = sVBCode
   
   If bStripContinuations Then sin = StripContinuations(sin)
   If bExpandColonCmds Then sin = ColonsToLineBreaks(sin)
   If bStandardize Then sin = StandardizeSpaces(sin)
   
   internal_IndentVBCode sin, sAllLines
   
   If bAddNewLines Then
        For j = o To UBound(sAllLines)
        
            addit = False
            If AnyOfTheseInstr(firstWord(sAllLines(j)), "for,if,select,do,while,else,case,call,function,private,sub,property") Then
                If j = 0 Then
                    addit = True
                Else
                    If Len(sAllLines(j - 1)) > 0 Then addit = True
                End If
                If addit And j > 0 Then
                    If sAllLines(j - 1) = vbCrLf Then addit = False
                End If
                If addit Then sAllLines(j) = vbCrLf & sAllLines(j)
            End If
            
            addit = False
            If AnyOfTheseInstr(firstWord(sAllLines(j)), "end,loop,next") Then
                If j = UBound(sAllLines) Then
                    addit = True
                Else
                    If Len(sAllLines(j + 1)) > 0 Then addit = True
                End If
                If addit And j < UBound(sAllLines) Then
                    If sAllLines(j + 1) = vbCrLf Then addit = False
                End If
                If addit Then sAllLines(j) = sAllLines(j) & vbCrLf
            End If
            
        Next
    End If
        
    'sAlllines elements have been modified and now contain embedded crlf in them
    'this is not redundant even though it looks like it at a glance..
    tmp = Join(sAllLines, vbCrLf)
    tmp = Replace(tmp, vbCrLf & vbCrLf, vbCrLf)
    tmp = Split(tmp, vbCrLf)
    
    sAllLines = tmp
    IndentVBCode = Join(tmp, vbCrLf)
    
End Function

Private Function internal_IndentVBCode(sVBCode As String, sAllLines() As String) As String
                
   Dim nI               As Long
   Dim nJ               As Long
   Dim nNumberLines     As Long
   Dim nTab             As Integer
   Dim sTmp             As String
   Dim nPos             As Integer
   Dim nIndent          As Byte
   Dim sTWA()           As String
   
   RaiseEvent Progress(0)
   sAllLines = Split(sVBCode, vbCrLf)
   nNumberLines = UBound(sAllLines)

   ReDim sBuffer(0 To nNumberLines)

   ' *** Determine number of tabs
   For nI = 0 To nNumberLines
      
      If nI Mod 50 = 0 Then
          RaiseEvent Progress(CInt((nI / nNumberLines) * 100))
      End If

      sAllLines(nI) = Replace(sAllLines(nI), Chr$(9), "")

      Select Case BlockType(sAllLines(nI))
         Case LABELLINE:
            sBuffer(nI).nNumTabs = 0

         Case STARTBLOCK:
            For nJ = nI + 1 To nNumberLines
               sBuffer(nJ).nNumTabs = sBuffer(nJ).nNumTabs + 1
            Next

         Case NOINDENTBLOCK:
            sBuffer(nI).nNumTabs = 0

         Case MIDBLOCK:
            nTab = sBuffer(nI).nNumTabs
            For nJ = nI + 1 To nNumberLines
               sBuffer(nJ).nNumTabs = nTab
            Next
            sBuffer(nI).nNumTabs = nTab - 1

         Case CASEBLOCK:
            sBuffer(nI).nNumTabs = sBuffer(nI).nNumTabs
            For nJ = nI + 1 To nNumberLines
               sBuffer(nJ).nNumTabs = sBuffer(nJ).nNumTabs + 1
            Next

         Case ENDCASEBLOCK:
            sBuffer(nI).nNumTabs = sBuffer(nI).nNumTabs - 2
            nTab = sBuffer(nI).nNumTabs
            For nJ = nI + 1 To nNumberLines
               sBuffer(nJ).nNumTabs = nTab
            Next

         Case ENDBLOCK:
            sBuffer(nI).nNumTabs = sBuffer(nI).nNumTabs - 1
            nTab = sBuffer(nI).nNumTabs
            For nJ = nI + 1 To nNumberLines
               sBuffer(nJ).nNumTabs = nTab
            Next

         Case NOALIGNCOMMENTLOCK:
            sBuffer(nI).nNumTabs = sBuffer(nI).nNumTabs - 1

         Case UNDERSCORECHAR:
            If nI + 1 <= nNumberLines Then sBuffer(nI + 1).nNumTabs = sBuffer(nI + 1).nNumTabs + 1

         Case VARIABLELINE:
            If bIndentVariable Then
               On Error Resume Next
               'vars.Add sAllLines(nI)
               nPos = InStr(LCase$(sAllLines(nI)), " as ")
               sTmp = Trim$(Left$(sAllLines(nI), nPos))
               If Left$(LCase$(sTmp), 3) = "dim" Then
                  If sBuffer(nI).nNumTabs = 0 Then
                     nIndent = 24
                  Else
                     nIndent = 21
                  End If
               ElseIf Left$(LCase$(sTmp), 6) = "global" Then
                  nIndent = 24
               ElseIf Left$(LCase$(sTmp), 6) = "public" Then
                  nIndent = 24
               ElseIf Left$(LCase$(sTmp), 7) = "private" Then
                  nIndent = 24
               End If

               If Len(sTmp) < nIndent Then
                  sAllLines(nI) = PadR(sTmp, CInt(nIndent)) & Mid$(sAllLines(nI), nPos + 1)
               Else
                  sAllLines(nI) = sTmp & " " & Mid$(sAllLines(nI), nPos + 1)
               End If
               On Error GoTo ERROR_IndentVBCode
            End If

         Case UNALTERED:
            If bIndentVariable Then
               On Error Resume Next
               If Right$(Trim$(sAllLines(nI)), 1) <> "_" Then
                  nPos = InStr(LCase$(sAllLines(nI)), " as ")
                  If nPos > 0 Then
                     sTmp = Trim$(sAllLines(nI))
                     sTmp = Replace(sTmp, "            ", " ")
                     sTmp = Replace(sTmp, "           ", " ")
                     sTmp = Replace(sTmp, "          ", " ")
                     sTmp = Replace(sTmp, "         ", " ")
                     sTmp = Replace(sTmp, "        ", " ")
                     sTmp = Replace(sTmp, "       ", " ")
                     sTmp = Replace(sTmp, "      ", " ")
                     sTmp = Replace(sTmp, "     ", " ")
                     sTmp = Replace(sTmp, "    ", " ")
                     sTmp = Replace(sTmp, "   ", " ")
                     sTmp = Replace(sTmp, "  ", " ")
                     sTmp = Replace(sTmp, ",", " ")
                     sTWA = Split(sTmp, " ")
                     If LCase$(sTWA(2)) = "as" Then
                        sTmp = Trim$(Left$(sAllLines(nI), nPos))
                        nIndent = 21

                        If Len(sTmp) < nIndent Then
                           sAllLines(nI) = PadR(sTmp, CInt(nIndent)) & Mid$(sAllLines(nI), nPos + 1)
                        Else
                           sAllLines(nI) = sTmp & " " & Mid$(sAllLines(nI), nPos + 1)
                        End If
                        On Error GoTo ERROR_IndentVBCode
                     End If
                  End If
               End If
            End If

      End Select
   Next

   If bUseTabs Then nIndentSpaces = 1

   On Error GoTo ERROR_IndentVBCode
   ' *** Add tabs in front of each line
   For nI = 1 To nNumberLines
      If sBuffer(nI).nNumTabs < 0 Then sBuffer(nI).nNumTabs = 0
      sAllLines(nI) = Trim$(sAllLines(nI))
      If Len(sAllLines(nI)) <> 0 Then
         ' *** Remove all tabs
         Do While Asc(Left$(sAllLines(nI), 1)) = 9
            sAllLines(nI) = Right$(sAllLines(nI), Len(sAllLines(nI)) - 1)
         Loop
         sTmp = LCase$(sAllLines(nI))
         If ((sTmp = "end sub") Or (sTmp = "end function") Or (sTmp = "end property")) And (nI < nNumberLines - 1) Then
            ' *** Verify if the next line is empty or not
            If Trim$(sAllLines(nI + 1)) <> "" Then sAllLines(nI) = sAllLines(nI) & vbCrLf
         End If

         sAllLines(nI) = String$(nIndentSpaces * sBuffer(nI).nNumTabs, sCharIndent) & sAllLines(nI) & vbCrLf
      Else
         sAllLines(nI) = Trim$(sAllLines(nI)) & vbCrLf
      End If
   Next

ERROR_IndentVBCode:
   sTmp = Join(sAllLines, vbCrLf)
   internal_IndentVBCode = Replace(sTmp, vbCrLf & vbCrLf & vbCrLf, vbCrLf & vbCrLf)
   RaiseEvent Progress(0)

End Function

Private Function BlockType(ByVal sLine As String) As Byte

   Dim sFirstWord       As String
   Dim sSecondWord      As String
   Dim sLastWord        As String
   Dim bUnderscore      As Boolean

   Dim nPos             As Long
   Dim nPos2            As Long

   Static bSelectCase   As Boolean
   Static nCase         As Integer
   Static nPrevBlock    As Integer

   ' *** Get the needed words
   Call GetWords(sLine, sFirstWord, sSecondWord, sLastWord)

   If sLastWord = "_" Then
      ' *** Middle Code
      If (sFirstWord = "else") Or (sFirstWord = "#else") Or (sFirstWord = "elseif") Then
         BlockType = MIDBLOCK
         nPrevBlock = BlockType
         gbUnderscore = True
         Exit Function
      End If

      If gbUnderscore = False Then
         If Len(gsFirstWord) = 0 Then gsFirstWord = sFirstWord
         If Len(gsSecondWord) = 0 Then gsSecondWord = sSecondWord
         gbUnderscore = True
         BlockType = UNDERSCORECHAR
      Else
         'If gsFirstWord = "if" Then
         BlockType = UNDERSCORECHAR
         'Else
         '   BlockType = UNALTERED
         'End If
      End If
      nPrevBlock = BlockType
      Exit Function
   End If

   If gbUnderscore Then
      If (sSecondWord <> "") And (gsFirstWord = "if") And (sLastWord = "") And (sFirstWord = "then") Then
         sFirstWord = gsFirstWord
         sLastWord = sSecondWord
         sSecondWord = gsSecondWord
      Else
         If Len(sLastWord) = 0 Then sLastWord = sFirstWord
         sFirstWord = gsFirstWord
         sSecondWord = gsSecondWord
      End If

      gbUnderscore = False
      gsFirstWord = ""
      gsSecondWord = ""
   End If

   BlockType = UNALTERED

   ' *** Special treatment for comments
   If ((Left$(Trim$(sLine), 1) = "'") Or _
      (LCase$(Left$(Trim$(sLine), 3)) = "rem")) And bIndentCmt = False Then

      BlockType = NOALIGNCOMMENTLOCK
      nPrevBlock = BlockType
      Exit Function
   End If

   ' *** Label line
   If (Right$(sLine, 1) = ":") And (Left$(Trim$(sLine), 1) <> "'") And (sFirstWord <> "case") And (sFirstWord <> "if") Then
      nPos = InStrRev(sLine, "'") ' *** Search Comment
      nPos2 = InStrRev(sLine, """") ' *** Search String
      If (nPos = 0) Or (nPos2 > nPos) Then
         BlockType = LABELLINE
         nPrevBlock = BlockType
         Exit Function
      End If
   End If

   ' *** Special treatment for Select/Case
   If sFirstWord = "case" Then
      If bIndentCase Then
         If bSelectCase = True Then
            bSelectCase = False
            BlockType = CASEBLOCK
         Else
            BlockType = MIDBLOCK
         End If
         nCase = nCase + 1
      Else
         BlockType = MIDBLOCK
      End If
      nPrevBlock = BlockType
      Exit Function
   End If

   ' *** Some unaltered code
   If sFirstWord = "if" And sLastWord <> "then" Then Exit Function

   ' *** Middle Code
   If (sFirstWord = "else") Or _
      (sFirstWord = "#else") Or _
      (sFirstWord = "elseif") Then
      BlockType = MIDBLOCK
      nPrevBlock = BlockType
      Exit Function
   End If

   ' *** Indent Code
   If (sFirstWord = "if") Or _
      (sFirstWord = "#if") Or _
      (sFirstWord = "type") Or _
      (sFirstWord = "enum") Or _
      (sFirstWord = "with") Or _
      (sFirstWord = "for") Or _
      (sFirstWord = "do") Or _
      (sFirstWord = "while") Or _
      ((sFirstWord = "select") And (sSecondWord = "case")) Or _
      ((sFirstWord = "public") And (sSecondWord = "type")) Or _
      ((sFirstWord = "private") And (sSecondWord = "type")) Or _
      ((sFirstWord = "public") And (sSecondWord = "enum")) Or _
      ((sFirstWord = "private") And (sSecondWord = "enum")) Then

      If ((sFirstWord = "select") And (sSecondWord = "case")) Then bSelectCase = True
      If ((sFirstWord = "type") And (sSecondWord = "as")) Then
         BlockType = UNALTERED
         Exit Function
      End If

      If (nPrevBlock <> STARTBLOCK) Or (sFirstWord <> "enum") Then
         BlockType = STARTBLOCK
         nPrevBlock = BlockType
         Exit Function
      End If
   End If

   ' *** Outdent Code
   If (sFirstWord = "next") Or _
      (sFirstWord = "loop") Or _
      (sFirstWord = "wend") Or _
      ((sFirstWord = "end") And (sSecondWord = "if")) Or _
      ((sFirstWord = "#end") And (sSecondWord = "if")) Or _
      ((sFirstWord = "end") And (sSecondWord = "select")) Or _
      ((sFirstWord = "end") And (sSecondWord = "type")) Or _
      ((sFirstWord = "end") And (sSecondWord = "enum")) Or _
      ((sFirstWord = "end") And (sSecondWord = "with")) Then

      If ((sFirstWord = "end") And (sSecondWord = "select")) And (nCase > 0) Then
         nCase = nCase - 1
         BlockType = ENDCASEBLOCK
         nPrevBlock = BlockType
         Exit Function
      End If

      BlockType = ENDBLOCK
      nPrevBlock = BlockType
      Exit Function
   End If

   If bIndentProc = False Then Exit Function

   ' *** Indent procedure
   If (sFirstWord = "sub") Or _
      (sFirstWord = "function") Or _
      (sFirstWord = "property") Or _
      (sFirstWord = "for") Or _
      (sFirstWord = "do") Or _
      (sFirstWord = "while") Or _
      ((sFirstWord = "static") And (sSecondWord = "sub")) Or _
      ((sFirstWord = "static") And (sSecondWord = "function")) Or _
      ((sFirstWord = "private") And (sSecondWord = "sub")) Or _
      ((sFirstWord = "private") And (sSecondWord = "function")) Or _
      ((sFirstWord = "private") And (sSecondWord = "property")) Or _
      ((sFirstWord = "private") And (sSecondWord = "static")) Or _
      ((sFirstWord = "friend") And (sSecondWord = "sub")) Or _
      ((sFirstWord = "friend") And (sSecondWord = "function")) Or _
      ((sFirstWord = "friend") And (sSecondWord = "property")) Or _
      ((sFirstWord = "friend") And (sSecondWord = "static")) Or _
      ((sFirstWord = "public") And (sSecondWord = "sub")) Or _
      ((sFirstWord = "public") And (sSecondWord = "function")) Or _
      ((sFirstWord = "public") And (sSecondWord = "property")) Or _
      ((sFirstWord = "public") And (sSecondWord = "static")) Then

      ' *** Procedure in 1 line
      If ((sFirstWord = "sub") And (sLastWord = "sub")) Or _
         ((sFirstWord = "function") And (sLastWord = "function")) Or _
         ((sFirstWord = "property") And (sLastWord = "property")) Or _
         ((sSecondWord = "sub") And (sLastWord = "sub")) Or _
         ((sSecondWord = "function") And (sLastWord = "function")) Or _
         ((sSecondWord = "property") And (sLastWord = "property")) Then
         nPrevBlock = BlockType
         Exit Function
      End If

      ' *** Indent procedure
      BlockType = STARTBLOCK
      nPrevBlock = BlockType
      Exit Function
   End If

   ' *** Outdent procedure
   If ((sFirstWord = "end") And (sSecondWord = "sub")) Or _
      ((sFirstWord = "end") And (sSecondWord = "function")) Or _
      ((sFirstWord = "end") And (sSecondWord = "property")) Then
      BlockType = ENDBLOCK
      nPrevBlock = BlockType
      Exit Function
   End If

   If (sFirstWord = "dim") Or _
      (sFirstWord = "public") Or _
      (sFirstWord = "private") Or _
      (sFirstWord = "global") Then

      nPos = InStr(LCase$(sLine), " as ")
      If nPos > 0 Then
         BlockType = VARIABLELINE
      End If

   End If

   nPrevBlock = BlockType

End Function

Private Sub GetWords(ByVal sLine As String, sFirstWord As String, sSecondWord As String, sLastWord As String)

   Dim nI               As Long
   Dim nJ               As Long

   sFirstWord = ""
   sSecondWord = ""
   sLastWord = ""

   sLine = Trim$(LCase$(sLine)) ' this line is not optim.

   ' *** Remove things in strings
   nI = InStr(1, sLine, """")
   Do Until nI = 0
      nJ = InStr(nI + 1, sLine, """")
      If nJ = 0 Then nJ = nI + 1
      sLine = Left$(sLine, nI) & Mid$(sLine, nJ)
      nI = InStr(nI + 2, sLine, """")
   Loop

   ' *** Remove trailing comments from the line
   nI = InStr(1, sLine, "'")
   If nI > 0 Then sLine = Left$(sLine, nI - 1)

   nI = InStr(sLine, Chr$(39)) - 1
   If nI > 0 Then sLine = Left$(sLine, nI)

   If Left$(sLine, 1) = Chr$(39) Then Exit Sub
   If Len(sLine) = 0 Then Exit Sub

   Dim nSize            As Integer

   sLine = Trim$(sLine)

   nSize = Len(sLine)
   If nSize = 0 Then Exit Sub
   
   ReDim byteArrayString(1 To nSize)

   ' *** Copy string to byte array
   Call CopyMemory(byteArrayString(1), ByVal sLine, nSize)

   Dim nStartWord       As Integer
   Dim nSavPosition     As Integer

   ' *** Get FirstWord
   nStartWord = 1
   For nI = 1 To nSize
      If isSeparator(byteArrayString(nI)) Then
         sFirstWord = Space$(nI - nStartWord)
         CopyMemory ByVal sFirstWord, byteArrayString(nStartWord), nI - nStartWord
         nSavPosition = nI + 1
         Exit For
      End If
   Next

   If nSavPosition = 0 Then
      sFirstWord = Space$(nI - nStartWord)
      CopyMemory ByVal sFirstWord, byteArrayString(nStartWord), nI - nStartWord
      Exit Sub
   End If

   ' *** Remove all blanks
   If nSavPosition < nSize Then
      Do While (byteArrayString(nSavPosition) = 32) And (nSavPosition < nSize)
         nSavPosition = nSavPosition + 1
      Loop
   End If

   ' *** Get SecondWord
   nStartWord = nSavPosition
   For nI = nSavPosition To nSize
      If isSeparator(byteArrayString(nI)) Then
         sSecondWord = Space$(nI - nStartWord)
         CopyMemory ByVal sSecondWord, byteArrayString(nStartWord), nI - nStartWord
         nSavPosition = nI + 1
         Exit For
      End If
   Next

   If nSavPosition = nStartWord Then
      sSecondWord = Space$(nI - nStartWord)
      If (nI - nStartWord) > 0 Then
         CopyMemory ByVal sSecondWord, byteArrayString(nStartWord), nI - nStartWord
      End If
      sLastWord = sSecondWord
      Exit Sub
   End If

   ' *** Get LastWord
   For nI = nSize To nSavPosition Step -1
      If isSeparator(byteArrayString(nI)) Then
         sLastWord = Space$(nSize - nI)
         If (nSize - nI) > 0 Then
            CopyMemory ByVal sLastWord, byteArrayString(nI + 1), nSize - nI
         End If
         Exit For
      End If
   Next

   If (nI = nSavPosition - 1) Then
      If isSeparator(byteArrayString(nI)) Then
         sLastWord = Space$(nSize - nI)
         If (nSize - nI) > 0 Then
            CopyMemory ByVal sLastWord, byteArrayString(nI + 1), nSize - nI
         End If
      End If
   End If

End Sub

Private Sub Class_Terminate()
   Erase sAllLines
End Sub


Public Function PadR(sTmp As String, lNbr As Long) As String
   PadR = Left$(RTrim$(sTmp) & Space$(lNbr), lNbr)
End Function

Public Function PadL(sTmp As String, lNbr As Long) As String
   PadL = Right$(Space$(lNbr) & RTrim$(sTmp), lNbr)
End Function
 


'--------------------------------------------------------------------------
'quick hacks below..
'--------------------------------------------------------------------------

Private Function ColonsToLineBreaks(sin) As String
    
    Dim inString As Boolean
    
    'if there is a colon in a comment..lets just strip it to be easy..
    tmp = Split(sin, vbCrLf)
    For i = 0 To UBound(tmp)
        Y = Trim(Replace(tmp(i), vbTab, ""))
        If Left(Y, 1) = "'" Then
            If InStr(Y, ":") > 0 Then
                'its a single full line comment with a colon in it..
                tmp(i) = Replace(tmp(i), ":", " ")
            End If
        End If
    Next
        
    tmp = Join(tmp, vbCrLf)
    tmp = Replace(tmp, """""", Chr(4)) 'replace double quotes as chr(4) (embedded double quotes and empty strings)
    tmp = Split(tmp, """")
    For i = 0 To UBound(tmp)
        If Not inString Then
            tmp(i) = HandleCodeBlock(tmp(i))
        End If
        inString = Not inString
    Next
        
    tmp = Join(tmp, """")
    If Right(tmp, 1) = """" Then tmp = Mid(tmp, 1, Len(tmp) - 1)
    tmp = Replace(tmp, Chr(4), """""")
    
    ColonsToLineBreaks = tmp
    
End Function

Private Function StandardizeSpaces(sin) As String
    
    Dim inString As Boolean
    
    tmp = Replace(sin, """""", Chr(4)) 'replace double quotes as chr(4) (embedded double quotes and empty strings)
    tmp = Split(tmp, """")
    
    For i = 0 To UBound(tmp)
        If Not inString Then
            tmp(i) = Replace(tmp(i), vbTab, " ")
            While InStr(tmp(i), "  ") > 0
                tmp(i) = Replace(tmp(i), "  ", " ")
            Wend
            'tmp(i) = Trim(tmp(i))
            tmp(i) = Replace(tmp(i), " , ", ",") 'in case only some
            tmp(i) = Replace(tmp(i), "( ", "(")
            tmp(i) = Replace(tmp(i), " = ", "=")
            
            tmp(i) = Replace(tmp(i), ",", " , ") 'now all
            tmp(i) = Replace(tmp(i), "(", "( ")
            tmp(i) = Replace(tmp(i), "=", " = ")
            While InStr(tmp(i), "  ") > 0
                tmp(i) = Replace(tmp(i), "  ", " ")
            Wend
        End If
        inString = Not inString
    Next
        
    tmp = Join(tmp, """")
    If Right(tmp, 1) = """" Then tmp = Mid(tmp, 1, Len(tmp) - 1)
    tmp = Replace(tmp, Chr(4), """""")
    
    StandardizeSpaces = tmp
    
End Function

Private Function HandleCodeBlock(sin) As String
    
    a = InStr(sin, ":")
    b = InStr(sin, "'")
        
    If a = 0 Then
        HandleCodeBlock = sin
        Exit Function
    End If
    
    If b = 0 Then
        HandleCodeBlock = Replace(sin, ":", vbCrLf)
        Exit Function
    End If
        
    'so now we have both an embedded quote and a colon to handle
    'and we know we are in a code section not a quoted string literal..
    If InStr(sin, vbCrLf) > 0 Then
        tmp = Split(sin, vbCrLf)
        For i = 0 To UBound(tmp)
            tmp(i) = ReplaceColonIfNotInComment(tmp(i))
        Next
        sin = Join(tmp, vbCrLf)
    Else
        sin = ReplaceColonIfNotInComment(sin)
    End If
        
    HandleCodeBlock = sin
    
End Function
    
'we know sin does not contain any string literal
'and this is only a single line of code..no crlf
Private Function ReplaceColonIfNotInComment(ByVal sin) As String

    Dim cmt As String, a As Long, b As Long
    
    a = InStr(sin, ":")
    b = InStr(sin, "'")
    
    If a > 0 Then
        If b > 0 Then
            cmt = Mid(sin, b - 1)
            sin = Mid(sin, 1, b - 1)
        End If
        sin = Replace(sin, ":", vbCrLf) & cmt
    End If
    
    ReplaceColonIfNotInComment = sin
    
End Function

Private Function StripContinuations(sin) As String
    
    Dim ret() As String, tmp() As String
    
    tmp = Split(sin, vbCrLf)
    For i = 0 To UBound(tmp)
        Y = Trim(tmp(i))
        While Right(Y, 1) = vbTab
            Y = Mid(Y, 1, Len(Y) - 1)
        Wend
        If Right(Y, 1) = "_" Then
            aryAdd ret, Mid(Y, 1, Len(Y) - 1)
        Else
            push ret, tmp(i)
        End If
    Next
    
    
    StripContinuations = Join(ret, vbCrLf)
    
End Function


Private Function aryAdd(ary, it)
    
    If AryIsEmpty(ary) Then
        push ary, it
    Else
        ary(UBound(ary)) = ary(UBound(ary)) & " " & it
    End If
    
End Function


Private Function firstWord(ByVal sLine) As String
    On Error Resume Next
    sLine = Trim(sLine)
    tmp = Split(sLine, " ")
    firstWord = tmp(0)
End Function




