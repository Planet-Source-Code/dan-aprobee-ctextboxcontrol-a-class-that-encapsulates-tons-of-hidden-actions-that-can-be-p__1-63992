VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CtxtControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type charrange
    cpMin As Long
    cpMax As Long
End Type

Private Type findtext
    chrg As charrange
    lpstrText As String
End Type

Private Type POINTAPI
   x As Long
   Y As Long
End Type

Public Enum enCase
    MATCH_CASE = &H4
    NO_MATCH_CASE = &H800
End Enum

Public Enum enWholeword
    WHOLE_WORD_ONLY = &H2
    NO_WHOLE_WORD = &H1000
End Enum

Private Const WM_USER As Long = &H400
Private Const EM_CHARFROMPOS As Long = &HD7
Private Const EM_EXLINEFROMCHAR As Long = (WM_USER + 54)
Private Const EM_FINDTEXT As Long = (WM_USER + 56)
Private Const EM_GETFIRSTVISIBLELINE As Long = &HCE
Private Const EM_GETLINECOUNT As Long = &HBA
Private Const EM_GETLINE As Long = &HC4
Private Const EM_LINEINDEX As Long = &HBB
Private Const EM_LINELENGTH As Long = &HC1
Private Const EM_LINESCROLL As Long = &HB6
Private Const EM_SETMARGINS As Long = &HD3
Private Const EC_LEFTMARGIN = &H1
Private Const EC_RIGHTMARGIN = &H2
  
Private Declare Function ScreenToClient Lib "user32.dll" (ByVal hwnd As Long, ByRef lpPoint As POINTAPI) As Long
Private Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTAPI) As Long
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lparam As Any) As Long
Private Declare Function SendMessageAsString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lparam As String) As Long
Private Declare Function SendMessageBynum& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lparam As Long)
 
'-----------------------------------------------
' this sub will scroll to the line number
' you specify
'-----------------------------------------------
Sub Scroll(Textbox As Control, NumberOfLines As Long)
  SendMessageBynum Textbox.hwnd, EM_LINESCROLL, _
                        0, NumberOfLines
End Sub
'-----------------------------------------------
' this sub will set the left and right margins
' for a textbox (regular or rich)
'-----------------------------------------------
Sub SetTextMargins(Textbox As Control, _
                      ByVal LeftMargin As Integer, _
                      ByVal RightMargin As Integer)
Dim lparam As Long
Dim s      As String
 
 With Textbox
   lparam = (LeftMargin + RightMargin * &H10000)
   'must use  SendMessageBynum  for this to work
   SendMessageBynum .hwnd, EM_SETMARGINS, _
                 EC_LEFTMARGIN Or EC_RIGHTMARGIN, _
                 lparam
             
   'reset the text (in regular textbox)
   'to make the right margin take
   If TypeOf Textbox Is Textbox Then
      s = .Text
      .Text = ""
      .Text = s
   End If
 End With
 
End Sub
'----------------------------------------------
' this function will return the line number the
' caret is currently at. (rich or regular)
'----------------------------------------------
Function CurrLineNumber(richtext As Control) As Long

 With richtext
    CurrLineNumber = SendMessage(.hwnd, _
                   EM_EXLINEFROMCHAR, 0, _
                   ByVal CLng(.SelStart))
 End With
 
End Function
'----------------------------------------------
' this function will return the line number of
' the first visible line of a rich textbox
'----------------------------------------------
Function TopVisibleLine(Textbox As Control)
  TopVisibleLine = (SendMessage(Textbox.hwnd, _
                  EM_GETFIRSTVISIBLELINE, 0, 0) + 1)
End Function
'----------------------------------------------
' this function will return the total number
' of lines in the textbox (rich or regular)
'----------------------------------------------
Function LineCount(Textbox As Control) As Long
  LineCount = SendMessage(Textbox.hwnd, _
                  EM_GETLINECOUNT, 0, 0)
End Function
'----------------------------------------------
' this function will return then length of
' a specified line (linenum)
'----------------------------------------------
Function LineLength(Textbox As Control, _
                       lineNum As Long) As Long
Dim lngFirstCharPos As Long

  lngFirstCharPos = SendMessage(Textbox.hwnd, _
                EM_LINEINDEX, lineNum, 0&)
  'length of line
  LineLength = SendMessage(Textbox.hwnd, _
               EM_LINELENGTH, lngFirstCharPos, 0&)
End Function
'----------------------------------------------
' this function will return (string) the contents
' of a specified line of a textbox (rich or regular)
'----------------------------------------------
Function LineContents(Textbox As Control, _
                      lineNumber As Long) As String
Dim lk      As String
Dim buffer  As String
Dim x       As Long
Dim byteLo  As Integer
Dim byteHi  As Integer

Const MAX_CHAR_PER_LINE = 2560
  
  byteLo = MAX_CHAR_PER_LINE And (255)
  byteHi = Int(MAX_CHAR_PER_LINE / 256)
  buffer = Chr(byteLo) + Chr(byteHi) + _
          Space(MAX_CHAR_PER_LINE - 2)
 
  x = SendMessageAsString(Textbox.hwnd, _
                  EM_GETLINE, ByVal lineNumber, _
                  ByVal buffer)
                  
  lk = Left(buffer, x)
  lk = LTrim(lk)
  LineContents = lk
              
End Function
'----------------------------------------------
' this function returns the character position
' of the word the user is searching for
' user has the option of forcing the search to
' match case or not and whether or not to search
' for whole word or not
'----------------------------------------------
Function Find(richtext As Control, strToFind As String, _
              searchStart As Long, searchEnd As Long, _
              MatchCase As enCase, _
              WholeWord As enWholeword) As Long
Dim fText As findtext
Dim cRange As charrange

 With cRange
   .cpMin = searchStart
   .cpMax = searchEnd
 End With

 With fText
   .chrg = cRange
   .lpstrText = strToFind
 End With
 
 Find = SendMessage(richtext.hwnd, _
        EM_FINDTEXT, MatchCase Or WholeWord, fText)
 
End Function
'----------------------------------------------
' this function will return the character(letter)
' under the mouse
'----------------------------------------------
Function ChrUnderMouse(richtext As Control) As String
Dim pos    As Long
 
 With richtext
   pos = funcChrPos(richtext)
   If pos > 0 Then
     ChrUnderMouse = Mid(.Text, pos, 1)
   End If
 End With
 
End Function
'----------------------------------------------
' this function will return the word that is
' under the mouse
'----------------------------------------------
Function WordUnderMouse(richtext As Control) As String
Dim pos    As Long
Dim i      As Integer
Dim lstart As Integer
Dim lend   As Integer
Dim llen   As Integer
Dim stxt   As String
Dim schr   As String

 With richtext
   llen = Len(.Text)
   stxt = .Text
   pos = funcChrPos(richtext)
   
   If pos > 0 Then
     'start from the chr position of the chr
     'under the mouse and work backwords one
     'character at a time looking for " " or
     'the first character in the textbox..in
     'other words, the beginning of the word
     'under the mouse
      For i = pos To 1 Step -1
        schr = Mid(stxt, i, 1)
        If schr = " " Or schr = Chr(13) Or i = 1 Then
          'if the starting character is vbcrlf then
          'we want to chop that off
          If schr = Chr(13) Then
            lstart = (i + 2)
          Else
            lstart = i
          End If
          Exit For
        End If
      Next i

     'start from the chr position of the chr
     'under the mouse and work fowards one
     'character at a time looking for " "
     'or the last character of the textbox...
     'in other words, the end of the word
      For i = pos To llen
        If Mid(stxt, i, 1) = " " Or _
           Mid(stxt, i, 1) = Chr(13) Or _
           i = llen Then
              lend = i
              Exit For
        End If
      Next i
      
      'prevent errors
      If lend >= lstart Then
          WordUnderMouse = Trim$(Mid$( _
               stxt, lstart, (lend - lstart)))
      End If
   End If
   
 End With
 
End Function

'we want this to be read only
Public Property Get FIND_MATCH_CASE() As Long
  FIND_MATCH_CASE = &H4
End Property
'we want this to be read only
Public Property Get FIND_NO_MATCH_CASE() As Long
  FIND_NO_MATCH_CASE = &H800
End Property
'we want this to be read only
Public Property Get FIND_WHOLE_WORD_ONLY() As Long
  FIND_WHOLE_WORD_ONLY = &H2
End Property
'we want this to be read only
Public Property Get FIND_NO_WHOLE_WORD_ONLY() As Long
  FIND_NO_WHOLE_WORD_ONLY = &H1000
End Property





'==============================================
'        PRIVATE SUBS AND FUNCTIONS
'==============================================




'----------------------------------------------
' this function will return the character
' position of the character the mouse is
' currently over
'----------------------------------------------
Private Function funcChrPos(Textbox As Control) As Long
Dim lhWnd  As Long
Dim pt As POINTAPI
 
 With Textbox
   lhWnd = .hwnd
   'where the cursor is
   GetCursorPos pt
   'convert coods relative to the textbox itself
   ScreenToClient lhWnd, pt
   funcChrPos = SendMessage(lhWnd, EM_CHARFROMPOS, 0&, pt)
 End With
 
End Function

 
