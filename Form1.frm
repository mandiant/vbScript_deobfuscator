VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "VbScript Deobsfuscator v0.1"
   ClientHeight    =   12075
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   14655
   LinkTopic       =   "Form1"
   ScaleHeight     =   12075
   ScaleWidth      =   14655
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSCOut 
      Height          =   330
      Left            =   13590
      TabIndex        =   18
      Top             =   45
      Visible         =   0   'False
      Width           =   735
   End
   Begin MSScriptControlCtl.ScriptControl sc 
      Left            =   12915
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      UseSafeSubset   =   -1  'True
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   285
      Left            =   11205
      TabIndex        =   17
      Top             =   90
      Width           =   600
   End
   Begin VB.TextBox txtLoadedPath 
      Height          =   285
      Left            =   5355
      OLEDropMode     =   1  'Manual
      TabIndex        =   16
      Top             =   90
      Width           =   5775
   End
   Begin VB.CommandButton cmdSaveRenamed 
      Caption         =   "Save"
      Height          =   285
      Left            =   13050
      TabIndex        =   14
      Top             =   6480
      Width           =   825
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Open"
      Height          =   285
      Left            =   11925
      TabIndex        =   13
      Top             =   90
      Width           =   600
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Indent / Parse"
      Height          =   420
      Left            =   6975
      TabIndex        =   9
      Top             =   495
      Width           =   1365
   End
   Begin VB.CommandButton Command2 
      Caption         =   "StartOver"
      Height          =   420
      Left            =   11970
      TabIndex        =   8
      Top             =   540
      Width           =   1320
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Test Tokenizer"
      Height          =   420
      Left            =   13365
      TabIndex        =   7
      Top             =   540
      Width           =   1230
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Rename"
      Height          =   420
      Left            =   8685
      TabIndex        =   6
      Top             =   495
      Width           =   1320
   End
   Begin MSComctlLib.ListView lv 
      Height          =   3120
      Left            =   45
      TabIndex        =   1
      Top             =   3105
      Width           =   4605
      _ExtentX        =   8123
      _ExtentY        =   5503
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Line"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Function Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Args"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "LVars"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "New Name"
         Object.Width           =   2540
      EndProperty
   End
   Begin RichTextLib.RichTextBox Text1 
      Height          =   5460
      Left            =   4770
      TabIndex        =   0
      Top             =   990
      Width           =   9825
      _ExtentX        =   17330
      _ExtentY        =   9631
      _Version        =   393217
      HideSelection   =   0   'False
      ScrollBars      =   3
      TextRTF         =   $"Form1.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ListView lv2 
      Height          =   2850
      Left            =   45
      TabIndex        =   2
      Top             =   6390
      Width           =   4560
      _ExtentX        =   8043
      _ExtentY        =   5027
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Func Args"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "newName"
         Object.Width           =   2540
      EndProperty
   End
   Begin RichTextLib.RichTextBox rtf2 
      Height          =   5100
      Left            =   4770
      TabIndex        =   3
      Top             =   6795
      Width           =   9870
      _ExtentX        =   17410
      _ExtentY        =   8996
      _Version        =   393217
      HideSelection   =   0   'False
      ScrollBars      =   3
      TextRTF         =   $"Form1.frx":0080
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ListView lv3 
      Height          =   2535
      Left            =   45
      TabIndex        =   4
      Top             =   450
      Width           =   4650
      _ExtentX        =   8202
      _ExtentY        =   4471
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Globals"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "newName"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lv4 
      Height          =   2535
      Left            =   90
      TabIndex        =   5
      Top             =   9360
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   4471
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Local Vars"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "newName"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "File"
      Height          =   240
      Left            =   4995
      TabIndex        =   15
      Top             =   135
      Width           =   375
   End
   Begin VB.Label Label6 
      Caption         =   "Renamed Text"
      Height          =   285
      Left            =   4815
      TabIndex        =   12
      Top             =   6480
      Width           =   1365
   End
   Begin VB.Label Label5 
      Caption         =   "Input Text"
      Height          =   285
      Left            =   4815
      TabIndex        =   11
      Top             =   675
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Double click any list to edit a new name, press d to keep  default"
      Height          =   285
      Left            =   90
      TabIndex        =   10
      Top             =   135
      Width           =   4740
   End
   Begin VB.Menu mnuTools 
      Caption         =   "Tools"
      Begin VB.Menu mnuCharStrings 
         Caption         =   "Char Strings"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LoadedFile As String
Dim test As String
Dim ci As New CIndenter
Dim rf As New CRefactor
Dim selLi As ListItem
Dim selLi2 As ListItem
Dim selLiGVar As ListItem
Dim selLiLVar As ListItem

Private Sub cmdBrowse_Click()
    Dim pth As String
    On Error Resume Next
    pth = dlg.OpenDialog(AllFiles)
    If Len(pth) = 0 Then Exit Sub
    txtLoadedPath = pth
    Me.Refresh
    cmdLoad_Click
End Sub

Private Sub cmdLoad_Click()
    Dim pth As String
    On Error Resume Next
    If Not FileExists(txtLoadedPath) Then
        cmdBrowse_Click
        Exit Sub
    End If
    Command2_Click
    LoadedFile = txtLoadedPath
    test = ReadFile(txtLoadedPath)
    Text1 = test
    modSyntaxHighlighting.SyntaxHighlight Text1
End Sub

Private Sub cmdSaveRenamed_Click()
    Dim pth As String, pdir As String
    On Error Resume Next
    pdir = GetParentFolder(LoadedFile)
    pth = dlg.SaveDialog(AllFiles, pdir, , , Me.hwnd, "renamed.txt")
    If Len(pth) = 0 Then Exit Sub
    WriteFile pth, rtf2.Text
    If FileExists(pth) Then MsgBox "Saved!", vbInformation
End Sub

Private Sub Command1_Click()
    
   
    Dim sTmp As String, tmp() As String
    
    StartBenchMark
    Set ci = New CIndenter
    sTmp = ci.IndentVBCode(Text1.Text, tmp())
    Text1.Text = Join(tmp, vbCrLf)
    Me.Caption = "Indent: " & EndBenchMark()
    
    Dim cf As CFunc
    Dim ca As CArg
    Dim li As ListItem
    
    lv.ListItems.Clear
    lv2.ListItems.Clear
    lv3.ListItems.Clear
    lv4.ListItems.Clear
    rtf2.Text = Empty
    
    StartBenchMark
    rf.Analyze sTmp
    Me.Caption = Me.Caption & " Parse: " & EndBenchMark()
    
    StartBenchMark
    rf.RenameAll
    Me.Caption = Me.Caption & " RenameAll: " & EndBenchMark()
    
    For Each cf In rf.methods
        Set li = lv.ListItems.Add(, , cf.LineNumber)
        li.SubItems(1) = cf.name
        li.SubItems(2) = cf.args.Count
        li.SubItems(3) = cf.localVars.Count
        li.SubItems(4) = cf.newName
        Set li.Tag = cf
    Next
    
    For Each ca In rf.gVars
        Set li = lv3.ListItems.Add(, , ca.name)
        li.SubItems(1) = ca.newName
        Set li.Tag = ca
    Next
 
    modSyntaxHighlighting.SyntaxHighlight Text1
    
End Sub

'test tokenizer...
Private Sub Command3_Click()
    Dim c As Collection
    Dim ct As CToken
    Dim ret() As String
    
    'Const t = "Set t = NextToken(x, pos)"
    'Const t = "Const sSeparators = vbTab & "" ,.:;!?""""()=-><+&"" & vbCrLf"
    'Const t = "Text1.Text = fso.ReadFile(App.Path & ""\sample.txt"")"
    'Const t = "a = ""test"""
    'Const t = "a = a & """
    t = "a=a-1"
    
    Set c = TokenizeLine(t)
    
    For Each ct In c
        Debug.Print ct.token & " | '" & ct.seperator & "'"
        push ret, ct.token & ct.seperator
    Next
    
    Text1 = t
    rtf2 = Join(ret, "")
    
    modSyntaxHighlighting.SyntaxHighlight Text1
    modSyntaxHighlighting.SyntaxHighlight rtf2
    
End Sub

Private Sub Command4_Click()
        
    If lv.ListItems.Count = 0 Then
        MsgBox "Run indent first"
        Exit Sub
    End If
    
    StartBenchMark
    'gs = rf.GlobalScript:rtf2.Text = gs: Exit Sub
    
    Dim cf As CFunc, i As Long, ca As CArg, j As Long, cf2 As CFunc
    
    'now lets go through each method and update the names..
    For Each cf In rf.methods
        
        cf.RebuildProto
        
        'global variables
        For Each ca In rf.gVars
            cf.RenameTokens ca.name, ca.newName
        Next
        
        'function arguments specific to this function..
        For Each ca In cf.args
            cf.RenameTokens ca.name, ca.newName
        Next
        
        'local variables specific to this function
        For Each ca In cf.localVars
            cf.RenameTokens ca.name, ca.newName
        Next
        
        'references to other functions..including recursive
        For Each cf2 In rf.methods
            cf.RenameTokens cf2.name, cf2.newName
        Next
        
    Next
    
    rtf2.Text = rf.RebuildGlobalScript
    Me.Caption = EndBenchMark
    
    modSyntaxHighlighting.SyntaxHighlight rtf2
        
End Sub

Private Sub lv_DblClick()
    Dim cf As CFunc
    If selLi Is Nothing Then Exit Sub
    Set cf = selLi.Tag
    
    newName = InputBox("Enter new name for " & cf.name, , cf.newName)
    If Len(newName) = 0 Then Exit Sub
    cf.newName = newName
    selLi.SubItems(4) = newName
    
End Sub

Private Sub lv2_DblClick()
    Dim ca As CArg
    If selLi2 Is Nothing Then Exit Sub
    Set ca = selLi2.Tag
    
    newName = InputBox("Enter new name for " & ca.name, , ca.newName)
    If Len(newName) = 0 Then Exit Sub
    ca.newName = newName
    selLi2.SubItems(1) = newName
    
End Sub

Private Sub lv3_DblClick()
    Dim ca As CArg
    If selLiGVar Is Nothing Then Exit Sub
    Set ca = selLiGVar.Tag
    
    newName = InputBox("Enter new name for " & ca.name, , ca.newName)
    If Len(newName) = 0 Then Exit Sub
    ca.newName = newName
    selLiGVar.SubItems(1) = newName
    
End Sub

Private Sub lv4_DblClick()
    Dim ca As CArg
    If selLiLVar Is Nothing Then Exit Sub
    Set ca = selLiLVar.Tag
    
    newName = InputBox("Enter new name for " & ca.name, , ca.newName)
    If Len(newName) = 0 Then Exit Sub
    ca.newName = newName
    selLiLVar.SubItems(1) = newName
    
End Sub

Private Sub lv_KeyPress(KeyAscii As Integer)
    
    Dim cf As CFunc
    If selLi Is Nothing Then Exit Sub
    
    If KeyAscii = Asc("d") Then
        Set cf = selLi.Tag
        newName = selLi.SubItems(1)
        If Len(newName) = 0 Then Exit Sub
        cf.newName = newName
        selLi.SubItems(4) = newName
        KeyAscii = 0
    End If
    
    If KeyAscii = Asc(" ") Then
        lv_DblClick
        KeyAscii = 0
    End If
    
End Sub

Private Sub lv2_KeyPress(KeyAscii As Integer)
    
    Dim cf As CArg
    If selLi2 Is Nothing Then Exit Sub
    
    If KeyAscii = Asc("d") Then
        Set cf = selLi2.Tag
        newName = selLi2.Text
        If Len(newName) = 0 Then Exit Sub
        cf.newName = newName
        selLi2.SubItems(1) = newName
        KeyAscii = 0
    End If
    
    If KeyAscii = Asc(" ") Then
        lv2_DblClick
        KeyAscii = 0
    End If
    
End Sub

Private Sub lv3_KeyPress(KeyAscii As Integer)
    
    Dim cf As CArg
    If selLiGVar Is Nothing Then Exit Sub
    
    If KeyAscii = Asc("d") Then
        Set cf = selLiGVar.Tag
        newName = selLiGVar.Text
        If Len(newName) = 0 Then Exit Sub
        cf.newName = newName
        selLiGVar.SubItems(1) = newName
        KeyAscii = 0
    End If
    
    If KeyAscii = Asc(" ") Then
        lv3_DblClick
        KeyAscii = 0
    End If
    
End Sub

Private Sub lv4_KeyPress(KeyAscii As Integer)
    
    Dim cf As CArg
    If selLiLVar Is Nothing Then Exit Sub
    
    If KeyAscii = Asc("d") Then
        Set cf = selLiLVar.Tag
        newName = selLiLVar.Text
        If Len(newName) = 0 Then Exit Sub
        cf.newName = newName
        selLiLVar.SubItems(1) = newName
        KeyAscii = 0
    End If
    
    If KeyAscii = Asc(" ") Then
        lv4_DblClick
        KeyAscii = 0
    End If
    
End Sub


Private Sub lv_ItemClick(ByVal Item As MSComctlLib.ListItem)

    Dim cf As CFunc, ca As CArg, li As ListItem
    
    Set selLi = Item
    Set cf = Item.Tag
    
    lv2.ListItems.Clear
    For Each ca In cf.args
        Set li = lv2.ListItems.Add(, , ca.name)
        li.SubItems(1) = ca.newName
        Set li.Tag = ca
    Next
    
    Set selLiLVar = Nothing
    lv4.ListItems.Clear
    For Each ca In cf.localVars
        Set li = lv4.ListItems.Add(, , ca.name)
        li.SubItems(1) = ca.newName
        Set li.Tag = ca
    Next
    
    Dim lLen As Long
    'first we go past it, then back to it so it shows up at top instead of bottom..
    GotoLine Text1, CInt(Item.Text) + 10
    GotoLine Text1, CInt(Item.Text), lLen
    Text1.selLength = lLen
    
    GotoLine rtf2, CInt(Item.Text) + 10
    GotoLine rtf2, CInt(Item.Text), lLen
    rtf2.selLength = lLen
    
End Sub

Private Sub lv2_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Set selLi2 = Item
End Sub

Private Sub lv3_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Set selLiGVar = Item
End Sub

Private Sub lv4_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Set selLiLVar = Item
End Sub


Private Sub Command2_Click()
    Text1 = test
    lv.ListItems.Clear
    lv2.ListItems.Clear
    lv3.ListItems.Clear
    lv4.ListItems.Clear
    rtf2.Text = Empty
    modSyntaxHighlighting.SyntaxHighlight Text1
End Sub

Private Sub Form_Load()
    On Error Resume Next
    
    Command3.Visible = IsIde()
    
    Text1.RightMargin = 1000000
    rtf2.RightMargin = 1000000
    
    Dim sample As String
    'sample = App.Path & "\test"
    'sample = App.Path & "\test2.txt"
    sample = App.path & "\sample.txt"
    
    If FileExists(sample) Then
        Text1.Text = ReadFile(sample)
        test = Text1.Text
        modSyntaxHighlighting.SyntaxHighlight Text1
    End If

End Sub
 
Private Sub mnuCharStrings_Click()
    Dim r As New RegExp
    Dim m As match
    Dim tmp As String
    Dim b() As Byte
    Dim i As Long
    Dim a As Byte
    Dim Z As Byte
    Dim zero As Byte
    Dim nine As Byte
    
    Dim isOk As Boolean
    Dim topLine As Long
    
    On Error Resume Next
    
    a = Asc("A")
    Z = Asc("Z")
    zero = Asc("0")
    nine = Asc("9")
    
    r.Global = True
    r.IgnoreCase = True
    r.Pattern = "(chr\([ &H0-9a-fA-F]+\)[ &\+\r\n]?)+" 'find all chr()+ strings built up (hex args ok)

    Set mm = r.Execute(Text1.Text)
    
    For Each m In mm
        
        
        tmp = vbsUnescapeChrString(m.value)
        
        If Len(tmp) > 0 Then
            b() = StrConv(UCase(tmp), vbFromUnicode, LANG_US)
            For i = 0 To UBound(b)
                isOk = False
                If b(i) = 20 Or b(i) = 9 Then isOk = True
                If b(i) >= a And b(i) <= Z Then isOk = True
                If b(i) >= zero And b(i) <= nine Then isOk = True
                If Not isOk Then
                    If AnyofTheseInstr2(Chr(b(i)), "~`!@#$%^&*()_+-={}[]|\;:<>,./?""'", vbBinaryCompare, "") Then isOk = True
                End If
                If Not isOk Then
                    Exit For
                End If
            Next
            If i = UBound(b) + 1 And isOk Then
                If InStr(tmp, """") > 0 Then tmp = Replace(tmp, """", """""") 'so we dont break js quoted strings..
                Text1.Text = Replace(Text1.Text, m.value, """" & tmp & """")
            End If
        End If
        
    Next
    
    
End Sub

Function vbsUnescapeChrString(str) As String
    
    On Error Resume Next
    Dim ret As String
    
    'txtSCOut.Text = str
    sc.Reset
    'sc.AddObject "txtSCOut", txtSCOut, True
    ret = sc.Eval(str)
    vbsUnescapeChrString = ret
    
End Function

Private Sub txtLoadedPath_OLEDragDrop(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single)
    On Error Resume Next
    If FileExists(data.Files(1)) Then
        txtLoadedPath = data.Files(1)
        cmdLoad_Click
    End If
End Sub
