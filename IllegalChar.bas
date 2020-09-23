Attribute VB_Name = "Module1"
Option Explicit

Private Function IsAlphaIntl(ByVal sChar As String) As Boolean

  'internationalised
  'safer than testing ASCII values becuase it also catches the high Ascii characters that
  'Anglos (including MS's IsAlpha API) never seem to remember/think about

  IsAlphaIntl = Not (UCase$(sChar) = LCase$(sChar))

End Function

Public Function IsPunctExcept(ByVal strTest As String, _
                              ByVal strExcept As String) As Boolean

  'Detect punctuation, but allows you to ignore any specified characters

  If InStr(strExcept, strTest) = 0 Then ' not marked to ignore
    If IsNumeric(strTest) Then          ' Is a number so allow
      IsPunctExcept = False
     Else
      IsPunctExcept = Not IsAlphaIntl(strTest) ' is a letter (or not)
    End If
  End If

End Function

Public Function RStrip(ByVal strInput As String, _
                       ByVal strStrip As String) As String

  'Strip strInput specified character from end of string
  'RTrim for rest of the characters

  If Right$(strInput, 1) = strStrip Then
    Do
      strInput = Left$(strInput, Len(strInput) - 1)
    Loop While Right$(strInput, 1) = strStrip
  End If
  RStrip = strInput

End Function

Public Function Strip(ByVal strInput As String, _
                      ByVal strStrip As String) As String

  'Strip strStrip from anywhere in StrInput

  Do While InStr(strInput, strStrip)
    strInput = Left$(strInput, InStr(strInput, strStrip) - 1) & Mid$(strInput, InStr(strInput, strStrip) + 1)
  Loop
  Strip = strInput

End Function

Public Function StripIllegalChars(ByVal strFix As String, _
                                  Optional strSubstitute As String = "_", _
                                  Optional ControlLength As Boolean = False) As String

  'OTIONAL PARAMETERS
  'strSubstitute replace illeagls with this value. routine can also jsut remove them if strSubstitute = ""
  'ControlLength True return string is <= 40 char in length False return is <= 255
  'you can then use Left$ to shorten the return further for sensibly sized names
  
  Dim I As Long

  I = 1
  StripIllegalChars = Trim$(StrConv(strFix, vbProperCase))
  'makes output more readable and gets rid of unwanted padding spaces
  Do While I <= Len(StripIllegalChars)
    'use 'Do' rather than 'For' because length of strFix may change)
    If IsPunctExcept(Mid$(StripIllegalChars, I, 1), "_") Then
      'replace/delete all unsafe characters
      If Len(strSubstitute) Then
        Mid$(StripIllegalChars, I, 1) = strSubstitute 'replace
       Else
        StripIllegalChars = Strip(StripIllegalChars, Mid$(StripIllegalChars, I, 1)) 'delete
        I = I - 1
      End If
    End If
    I = I + 1
  Loop
  'control/variable names cannot start with '_",spaces or numerals, so strip them off
  Do While IsNumeric(Left$(StripIllegalChars, 1)) Or InStr("_ ", Left$(StripIllegalChars, 1))
    StripIllegalChars = Mid$(StripIllegalChars, 2)
  Loop
  'just because end underscores look daggy
  StripIllegalChars = RStrip(StripIllegalChars, "_")
  If Len(strSubstitute) Then
    'take out multiple groups of strSubstitute
    'needs Do wrapper because of way Replace searchs for multiple instances
    ''Assume target string ="123___7890" (3 underscores at 4,5&6) and find is '__'(2 underscores)
    ''Replace will find char 4 & 5 and replace them with 1 underscore "123__7890"
    ''Replace will then resume searching at character 6 (found location + len(find)
    ''which is '7' so miss the new double char it created
    Do While InStr(StripIllegalChars, strSubstitute & strSubstitute)
      StripIllegalChars = Replace(StripIllegalChars, strSubstitute & strSubstitute, strSubstitute)
    Loop
  End If
  If Len(StripIllegalChars) > IIf(ControlLength, 40, 255) Then
    'controls cannot be longer than 40 (variable/filename(without extention) limit is 255, but don't do it)
    StripIllegalChars = Left$(StripIllegalChars, IIf(ControlLength, 40, 255))
  End If

End Function

''
''Public Function LStrip(ByVal strInput As String, ByVal strStrip As String) As String
''
'''Strip strInput specified character from start of string
'''LTrim for rest of the characters
''If Left$(strInput, 1) = strStrip Then
''Do
''strInput = Mid$(strInput, 2)
''Loop While Left$(strInput, 1) = strStrip
''End If
''LStrip = strInput
''End Function
''

':)Roja's VB Code Fixer V1.1.49 (8/11/2003 3:24:40 PM) 1 + 115 = 116 Lines Thanks Ulli for inspiration and lots of code.

