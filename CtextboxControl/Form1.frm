VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4125
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3615
   LinkTopic       =   "Form1"
   ScaleHeight     =   4125
   ScaleWidth      =   3615
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboScroll 
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   2385
      List            =   "Form1.frx":0016
      TabIndex        =   20
      Text            =   "2"
      Top             =   2205
      Width           =   600
   End
   Begin VB.CommandButton cmdScroll 
      Caption         =   "&Scroll..."
      Height          =   285
      Left            =   45
      TabIndex        =   18
      Top             =   2205
      Width           =   1005
   End
   Begin VB.ComboBox cboRightMargin 
      Height          =   315
      ItemData        =   "Form1.frx":002C
      Left            =   2745
      List            =   "Form1.frx":003C
      TabIndex        =   15
      Text            =   "10"
      Top             =   1890
      Width           =   645
   End
   Begin VB.ComboBox cboLeftMargin 
      Height          =   315
      ItemData        =   "Form1.frx":0050
      Left            =   1530
      List            =   "Form1.frx":0060
      TabIndex        =   14
      Text            =   "10"
      Top             =   1890
      Width           =   600
   End
   Begin VB.CommandButton cmdSetMargins 
      Caption         =   "&Set Margins"
      Height          =   285
      Left            =   45
      TabIndex        =   13
      Top             =   1890
      Width           =   1005
   End
   Begin VB.CheckBox ckWholeWord 
      Caption         =   "&whole word"
      Height          =   195
      Left            =   2250
      TabIndex        =   9
      Top             =   405
      Width           =   1230
   End
   Begin VB.CheckBox ckMatchCase 
      Caption         =   "&match case"
      Height          =   195
      Left            =   2250
      TabIndex        =   7
      Top             =   225
      Width           =   1230
   End
   Begin VB.TextBox txtFind 
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   810
      TabIndex        =   5
      Text            =   "Enter word to find"
      Top             =   270
      Width           =   1410
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Find"
      Height          =   285
      Left            =   45
      TabIndex        =   4
      Top             =   270
      Width           =   780
   End
   Begin RichTextLib.RichTextBox Rich 
      Height          =   1275
      Left            =   45
      TabIndex        =   0
      Top             =   585
      Width           =   3525
      _ExtentX        =   6218
      _ExtentY        =   2249
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"Form1.frx":0074
   End
   Begin VB.Label Label3 
      Caption         =   "to line number..."
      Height          =   240
      Left            =   1125
      TabIndex        =   19
      Top             =   2250
      Width           =   1275
   End
   Begin VB.Label Label2 
      Caption         =   "RIGHT"
      Height          =   195
      Left            =   2205
      TabIndex        =   17
      Top             =   1935
      Width           =   555
   End
   Begin VB.Label Label1 
      Caption         =   "LEFT"
      Height          =   195
      Left            =   1125
      TabIndex        =   16
      Top             =   1935
      Width           =   555
   End
   Begin VB.Label lblLineLen 
      Height          =   240
      Left            =   45
      TabIndex        =   12
      Top             =   3870
      Width           =   4110
   End
   Begin VB.Label lblLineContents 
      Height          =   240
      Left            =   45
      TabIndex        =   11
      Top             =   3645
      Width           =   4110
   End
   Begin VB.Label lblLineCount 
      Height          =   240
      Left            =   45
      TabIndex        =   10
      Top             =   3420
      Width           =   4110
   End
   Begin VB.Label lblTopLine 
      Height          =   240
      Left            =   45
      TabIndex        =   8
      Top             =   3195
      Width           =   4110
   End
   Begin VB.Label lblFind 
      Height          =   240
      Left            =   45
      TabIndex        =   6
      Top             =   0
      Width           =   3435
   End
   Begin VB.Label lblRichLine 
      Height          =   240
      Left            =   45
      TabIndex        =   3
      Top             =   2970
      Width           =   4110
   End
   Begin VB.Label lblRichWord 
      Height          =   240
      Left            =   45
      TabIndex        =   2
      Top             =   2745
      Width           =   4110
   End
   Begin VB.Label lblRichChr 
      Height          =   240
      Left            =   45
      TabIndex        =   1
      Top             =   2565
      Width           =   4110
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const CR = vbCrLf
Private cTxt      As CtxtControl
 
 

Private Sub cmdScroll_Click()
   cTxt.Scroll Rich, Val(cboScroll)
End Sub

Private Sub cmdSetMargins_Click()
  cTxt.SetTextMargins Rich, Val(cboLeftMargin.Text), _
                    Val(cboRightMargin.Text)
End Sub

Private Sub Form_Load()
Dim strtext As String
   
   strtext = "We must add some text to" & CR & _
             "this textbox so we can see the" & CR & _
             "results of using our" & CR & _
             "CtxtControl class!" & CR & _
             "The richtextbox is a powerful" & CR & _
             "control from microsoft but some" & CR & _
             "of its power is still hidden." & CR & _
             "Well this class unlocks some" & CR & _
             "more of that power!" & CR & _
             "So lets have some fun here" & CR & _
             "and dont forget to vote :)"
   Rich.Text = strtext
 
   Set cTxt = New CtxtControl
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Set cTxt = Nothing
End Sub

Private Sub cmdFind_Click()
Dim lfound   As Long
Dim lcase    As Long
Dim lwhole   As Long
  
With cTxt
  'are we matching case .. or not?
  If ckMatchCase.Value = vbChecked Then
    lcase = .FIND_MATCH_CASE
  Else
    lcase = .FIND_NO_MATCH_CASE
  End If
  'are we searching for whole words only?
  If ckWholeWord.Value = vbChecked Then
    lwhole = .FIND_WHOLE_WORD_ONLY
  Else
    lwhole = .FIND_NO_WHOLE_WORD_ONLY
  End If
  
  
  lfound = cTxt.Find(Rich, txtFind, 0, _
             Len(Rich), lcase, lwhole)
 
  ' a return of -1 means a match not found
  If lfound = -1 Then
    lblFind = "word not found"
  Else
    lblFind = "word found at position #: " & lfound
    'place cursor at beginning of the word
    Rich.SetFocus
    Rich.SelStart = lfound
  End If
End With
 
End Sub
 

Private Sub Rich_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  lblRichChr = "chr under mouse: " & cTxt.ChrUnderMouse(Rich)
  lblRichWord = "word under mouse: " & cTxt.WordUnderMouse(Rich)
  lblTopLine = "top visible line #: " & cTxt.TopVisibleLine(Rich)
End Sub

'------------------------------------------
' this event is raised anytime the carets
' position changes in the richtext for any
' reason
'------------------------------------------
Private Sub Rich_SelChange()
Dim currline As Long
  
  currline = cTxt.CurrLineNumber(Rich)
  lblRichLine = "The caret is at line #: " & _
                 (currline + 1)
  lblLineCount = "Total # lines: " & _
                  cTxt.LineCount(Rich)
  lblLineContents = "Current line: " & _
                  cTxt.LineContents(Rich, currline)
  lblLineLen = "Length of current line: " & _
                  cTxt.LineLength(Rich, currline)
End Sub
 

Private Sub Text1_Click()
   Caption = cTxt.LineLength(Text1, 2)
End Sub

Private Sub txtFind_GotFocus()
   If Trim$(txtFind) = "Enter word to find" Then
      txtFind = ""
   End If
End Sub

Private Sub txtFind_LostFocus()
  If Len(Trim$(txtFind)) = 0 Then
    txtFind = "Enter word to find"
  End If
End Sub

 
