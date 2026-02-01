Attribute VB_Name = "RICH"

Global lastgroup

Global ServerSentCount


Public Declare Function GetPrivateProfileString Lib "Kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "Kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Public Const SND_ASYNC = &H1

'the following constants are a prebuilt rtf header. i added the fonts and colors to be
'used in this example. i chose this method to make updating the rich edit controls fast
'and easy
Public Const RTF_HEADER As String = "{\rtf1\ansi\deff0\deftab720"
Public Const RTF_FONT_TABLE As String = "{\fonttbl{\f0\fswiss Arial;}}"
Public Const RTF_COLOR_TABLE As String = "{\colortbl\red0\green0\blue0;\red0\green0\blue255;\red255\green0\blue0;\red0\green130\blue0;\red0\green0\blue130;}"
Public Const RTF_START_TEXT As String = "\viewkind4\uc1\pard\cf1\lang1033"
Public Const RTF_START As String = RTF_HEADER & vbCrLf & RTF_FONT_TABLE & vbCrLf & RTF_COLOR_TABLE & vbCrLf & RTF_START_TEXT

Public m_lngLocalSeq As Long      'local sequence
Public m_lngServerSeq As Long     'incoming sequence (not really important)
Public m_strScreenName As String  'lower case screen name with no spaces
Public m_strPassword As String    'account password
Public m_strFormattedSN As String 'screen name formatted by server for display
Public m_strMode As String        'permit/deny mode
Public m_strPDList As String      'permit/deny list

Public strSoundSignOn As String   'buddy in sound
Public strSoundSignOff As String  'buddy out sound
Public strSoundFirstIM As String  'first im sound
Public strSoundIMIn As String     'im in sound
Public strSoundIMOut As String    'im out sound

Public strInviteBuddies As String 'list of buddies to send chat invite to
Public strInviteRoom As String    'chat room name for invite (not id)
Public strInviteMessage As String 'message to send with invite

Public MeAway As Boolean



'the following three procedures are used to help create the sequence numbers for the
'flap headers being sent to the aim toc server.
Public Function MakeLong(lngHi As Long, lngLo As Long) As Long
  MakeLong& = lngLo& * 256 + lngHi&
End Function

Public Function Lo(lngVal As Long) As Long
  Lo& = Fix(lngVal& / 256)
End Function

Public Function Hi(lngVal As Long) As Long
  Hi& = lngVal& Mod 256
End Function



Public Function KillHTML(ByVal strIn As String) As String
  'for the sake of this example, i chose not to try converting html to rtf. this method
  'is not perfect. should this have been a real client, i would have chosen to convert
  'the html to rtf. however, for the sake of this example, i chose just to remove as much
  'html as i could.
  Dim lngLen As Long, lngFound As Long, lngEnd As Long
  Dim strLeft As String, strRight As String
  strIn$ = Replace(strIn$, "<HTML>", "")
  strIn$ = Replace(strIn$, "</HTML>", "")
  strIn$ = Replace(strIn$, "<SUP>", "")
  strIn$ = Replace(strIn$, "</SUP>", "")
  strIn$ = Replace(strIn$, "<HR>", "")
  strIn$ = Replace(strIn$, "<H1>", "")
  strIn$ = Replace(strIn$, "<H2>", "")
  strIn$ = Replace(strIn$, "<H3>", "")
  strIn$ = Replace(strIn$, "<PRE>", "")
  strIn$ = Replace(strIn$, "</PRE>", "")
  strIn$ = Replace(strIn$, "<PRE=", "")
  strIn$ = Replace(strIn$, "<B>", "")
  strIn$ = Replace(strIn$, "</B>", "")
  strIn$ = Replace(strIn$, "<U>", "")
  strIn$ = Replace(strIn$, "</U>", "")
  strIn$ = Replace(strIn$, "<I>", "")
  strIn$ = Replace(strIn$, "</I>", "")
  strIn$ = Replace(strIn$, "<FONT>", "")
  strIn$ = Replace(strIn$, "</FONT>", "")
  strIn$ = Replace(strIn$, "<BODY>", "")
  strIn$ = Replace(strIn$, "</BODY>", "")
  strIn$ = Replace(strIn$, "<BR>", "")
  strIn$ = Replace(strIn$, "</A>", "")
  lngLen& = Len(strIn$)
  lngFound& = InStr(strIn$, "<BODY ")
  Do While lngFound& <> 0
    lngEnd& = InStr(lngFound&, strIn$, ">")
    If lngEnd& <> 0 Then
      strLeft$ = Left(strIn$, lngFound& - 1)
      strRight$ = Right(strIn$, lngLen& - lngEnd&)
      strIn$ = strLeft$ & strRight$
      lngLen& = Len(strIn$)
    End If
    lngFound& = InStr(lngFound& + 1, strIn$, "<BODY ")
  Loop
  lngFound& = InStr(strIn$, "<A ")
  Do While lngFound& <> 0
    lngEnd& = InStr(lngFound&, strIn$, ">")
    If lngEnd& <> 0 Then
      strLeft$ = Left(strIn$, lngFound& - 1)
      strRight$ = Right(strIn$, lngLen& - lngEnd&)
      strIn$ = strLeft$ & strRight$
      lngLen& = Len(strIn$)
    End If
    lngFound& = InStr(lngFound& + 1, strIn$, "<A ")
  Loop
  lngFound& = InStr(strIn$, "<FONT ")
  Do While lngFound& <> 0
    lngEnd& = InStr(lngFound&, strIn$, ">")
    If lngEnd& <> 0 Then
      strLeft$ = Left(strIn$, lngFound& - 1)
      strRight$ = Right(strIn$, lngLen& - lngEnd&)
      strIn$ = strLeft$ & strRight$
      lngLen& = Len(strIn$)
    End If
    lngFound& = InStr(lngFound& + 1, strIn$, "<FONT ")
  Loop
  strIn$ = Replace(strIn$, "&amp;", "&")
  strIn$ = Replace(strIn$, "&lt;", "<")
  KillHTML$ = strIn$
End Function

Public Function Normalize(ByVal strIn As String) As String
  'most strings sent to the aim toc server need to be normalized. this procedure formats
  'the strings as necessary.
  strIn$ = Replace(strIn$, "\", "\\")
  strIn$ = Replace(strIn$, "$", "$")
  strIn$ = Replace(strIn$, Chr(34), "\" & Chr(34))
  strIn$ = Replace(strIn$, "(", "\(")
  strIn$ = Replace(strIn$, ")", "\)")
  strIn$ = Replace(strIn$, "[", "\[")
  strIn$ = Replace(strIn$, "]", "\]")
  strIn$ = Replace(strIn$, "{", "\{")
  strIn$ = Replace(strIn$, "}", "\}")
  Normalize$ = strIn$
End Function


Public Function FixRTF(ByVal strRTF As String) As String
  'since we are updating with rtf, it is important to format some of our strings in order
  'to keep our rich text from showing up as rtf code.
  strRTF$ = Replace(strRTF$, "\", "\\")
  strRTF$ = Replace(strRTF$, "}", "\}")
  strRTF$ = Replace(strRTF$, "{", "\{")
  FixRTF$ = strRTF$
End Function

Public Function WriteINIString(strSection As String, strKeyName As String, strValue As String, strFile As String) As Long
  Dim lngStatus As Long
  lngStatus& = WritePrivateProfileString(strSection, strKeyName, strValue, strFile)
  WriteINIString& = (lngStatus& <> 0)
End Function

Public Function GetINIString(strSection As String, strKeyName As String, strFile As String, Optional strDefault As String = "") As String
  Dim strBuffer As String * 256, lngSize As Long
  lngSize& = GetPrivateProfileString(strSection$, strKeyName$, strDefault$, strBuffer$, 256, strFile$)
  GetINIString$ = Left$(strBuffer$, lngSize&)
End Function

Public Function FileExists(strFilename As String) As Boolean
  Dim intLen As Integer
  If strFilename$ <> "" Then
    intLen% = Len(Dir$(strFilename$))
    FileExists = (Not err And intLen% > 0)
  Else
    FileExists = False
  End If
End Function



Public Function FormByCaption(strMatch As String) As Long
  'since we are creating many forms dynamically, it is important for us to locate specific
  'forms. this procedure searches by the caption property while the one below searches
  'by the tag
  Dim lngDo As Long, lngFound As Long
  lngFound& = -1
  For lngDo& = 0 To Forms.Count - 1
    If LCase(Replace(Forms(lngDo&).Caption, " ", "")) = strMatch$ Then
      lngFound& = lngDo&
      Exit For
    End If
  Next
  FormByCaption& = lngFound&
End Function

Public Function FormByTag(strMatch As String) As Long
  Dim lngDo As Long, lngFound As Long
  lngFound& = -1
  For lngDo& = 0 To Forms.Count - 1
    If Forms(lngDo&).Tag = strMatch$ Then
      lngFound& = lngDo&
      Exit For
    End If
  Next
  FormByTag& = lngFound&
End Function

Public Function HideForm()
Dim lngDo As Long, lngFound As Long
  lngFound& = -1
  For lngDo& = 0 To Forms.Count - 1
    If MatchShit("*- Instant Message", Forms(lngDo&).Caption) = True Then
        'MsgBox "Found! - HideForm"
        Forms(lngDo&).Hide
    End If
  Next
  'FormByTag& = lngFound&
End Function

Public Function ShowForm()
Dim lngDo As Long, lngFound As Long
  lngFound& = -1
  For lngDo& = 0 To Forms.Count - 1
    If MatchShit("*- Instant Message", Forms(lngDo&).Caption) = True Then
    'MsgBox "Found! - ShowForm"
        Forms(lngDo&).Show
    End If
  Next
  'FormByTag& = lngFound&
End Function

Public Sub PlayWav(strWav As String)
  If FileExists(strWav$) Then Call sndPlaySound(strWav$, SND_ASYNC)
End Sub


Public Function MatchShit(a As String, b As String) As Boolean
Dim x
Dim ToFind As String
Dim start As Long
    'wild card text matcher so far it only s
    '
    ' upports one * but I'll
    'be trying to get it to work for unlimit
    '
    ' ed numbers of *, i havent found any
    'bugs. I saw a few subs that did somethi
    '
    ' ng like this but they were not as
    'good this is the first of it's kind, as
    '
    ' far as i know..
    '
    'if you can get it to work with unlimite
    '
    ' d number of *'s please email me it!
    'use the freely sub as you please long a
    '
    ' s you give me credit where it's due,
    'thanks! also dont forget to vote :)
    '
    'Using the sub: a is the string that is
    '
    ' the string that has an "*" in it
    'b is the base string meaing a is trying
    '
    ' to be matched to b.
    '
    'ie: a = 127.0.*, b = 127.0.0.1
    '
    'usage: MsgBox IPmatch(Text1, Text2)
    'usage2: if IPmatch(Text1, Text2) = fals
    '
    ' e then : msgbox "error no match!",16,"
    '     er
    ' ror"
    '
    '
    'Contacts:
    'Email: eko2k@hotmail.com
    'Aim:duckee2k
    '
    'oh and, btw open source 4lyfe
    On Error GoTo err
    'check the easy way :)
    If a = b Then: MatchShit = True: Exit Function
    'strings arent identical so here we go :
    '
    ' \
    x = Split(a, "*") 'searches For the *'s
    'defines the vars
    ToFind$ = Mid(a, InStr(a, x(1)))
    start& = InStr(a, x(1))
    'checks the legnth, 0,1,2+ need differnt
    '
    ' subs
    If Len(x(1)) = 1 Then GoTo short
    If Len(x(1)) = 0 Then GoTo none
    'core coding:
    'has more then one chr(s) after the *


    If Left(a, InStr(a, x(1)) - 2) = Left(b, InStr(a, x(1)) - 2) Then
        On Error GoTo err


        If Mid(a, InStr(a, x(1))) = Mid(b, InStr(start&, b, ToFind$)) Then
            MatchShit = True
        Else: GoTo err
        End If
    End If
    'has only one chr(s) after the *
short:


    If Left(a, InStr(a, x(1))) = Left(b, InStr(a, x(1))) Then
        On Error GoTo err


        If Right(a, InStr(a, x(1))) = Left(b, InStr(a, x(1))) Then
            MatchShit = True
        End If
    End If
    'has nothing after the *
none:


    If Len(a) = Len(b) Then
        On Error GoTo err


        If Left(a, Len(a) - 1) = Left(b, Len(a) - 1) Then
            MatchShit = True
        End If
    Else


        If Left(a, Len(a) - 1) = Left(b, Len(a) - 1) Then
            MatchShit = True
        End If
    End If
    Exit Function
err:
    MatchShit = False: Exit Function
End Function

Sub Sleep(ByVal MillaSec As Long, Optional ByVal DeepSleep As Boolean = False)
    Dim tStart#, Tmr#
    tStart = Timer


    While Tmr < (MillaSec / 1000)
        Tmr = Timer - tStart
        If DeepSleep = False Then DoEvents
    Wend
End Sub

 
Public Sub RTFUpdate(rtfOut As RichTextBox, strUpdate As String)
  Dim strRTF As String
  strRTF$ = RTF_START & strUpdate$ & "}"
  rtfOut.SelStart = Len(rtfOut.Text)
  rtfOut.SelRTF = strRTF$

  
  
  If Asc(Left$(rtfOut.Text, 1)) = 13 Then
  rtfOut.SelStart = 1
  rtfOut.SelLength = 2
  rtfOut.SelText = ""
  
  End If
  
End Sub
Public Function removeHTML(theStrin) 'remove everything in <>
Dim inCommand As Boolean
removeHTML = ""
For i = 1 To Len(theStrin)
    If inCommand = False Then
        If Mid(theStrin, i, 1) = "<" Then
            inCommand = True
            GoTo Skip:
        Else
            removeHTML = removeHTML & Mid(theStrin, i, 1)
        End If
    Else
        If Mid(theStrin, i, 1) = ">" Then
            inCommand = False
        End If
    End If
Skip:
    DoEvents
Next i
End Function
