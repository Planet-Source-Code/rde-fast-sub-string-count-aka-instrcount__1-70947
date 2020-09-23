Attribute VB_Name = "mSubStrCnt"
Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''©Rd'
'
'                        Sub-String Count
'
'  This function searches the passed string for occurences of the
'  specified sub-string. It has the ability to make case-sensitive
'  or case-insensitive searches.
'
'  It can also perform whole-word-only matching using a unique
'  delimiter function included in this module.
'
'  SubStrCount will return the number of matches or zero if none.
'
'                     Extended Functionality
'
'  This SubStrCount implementation offers extended functionality
'  through the use of the optional lHitLimit parameter. This allows
'  it to be used in a similar way to other token style functions.
'
'  By passing the lHitLimit parameter as any positive value allows
'  you to limit how many matches are found in the current call, and
'  the value of the lStartPos parameter *is modified* to identify the
'  start position in the search string of the last sub-string found
'  (or zero if none found).
'
'  In this case, the function will return a value equal to or less
'  than the value of the lHitLimit parameter, and zero if none found.
'
'  Using this feature you can limit the number of matches found, and
'  make subsequent calls to SubStrCount by passing lStartPos + 1
'  (or lStartPos + Len(sSubStr)) to step through the search process
'  as needed, and stop when the function returns zero.
'
'                        Whole-Word-Only
'
'  By default, all non-alphabetic characters (with the exception of
'  underscores) are automatically treated as word delimiters when
'  performing whole-word-only seaches and do not need to be specified.
'
'  As only alphabetic characters are treated as non-delimiters you
'  can specify custom non-delimiters, that is, any character(s) can
'  be specified as part of whole words and therefore be treated as
'  non-delimiters.
'
'  To make numerical characters part of whole words and so set
'  as non-delimiters *by default* add this line to the IsDelim
'  function's select case statement:
'      Case 48 To 57: IsDelim = False
'
'  To specify custom/run-time changes to the list of delimiters make
'  a call to the public SetDelim subroutine and add character(s) to
'  be handled as part of whole words (or as delimiters):
'      SetDelim "1234567890", False
'
'  Remember, all non-alphabetic characters are already treated as
'  word delimiters and so do not need to be specified through a
'  call to SetDelim ???, True. Also, alphabetic characters can be
'  treated as word delimiters through a call to SetDelim "a", True.
'
'  Most delimiter implementations build a list/array to hold all
'  delimiters, but this modules approach is *much* faster.
'
'                            Notes
'
'  Passing lStartPos with a value < 1 will not cause an error; it
'  will default to 1 and start the search at the first character in
'  the search string.
'
'  The lStartPos parameter will be reset appropriately if lHitLimit
'  is specified > zero, but will be *left unchanged* if lHitLimit
'  is omitted or passed with a value <= zero.
'
'                          Free Usage
'
'  You may use this code in any way you wish, but please respect
'  my copyright. But, if you can modify this function in some way
'  to speed it up or to add extra features then you can claim it
'  as your own!
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private mDelimList As String
Private mNotDelim As String

'Sub-String Count''''''''''''''''''''''''''''''''''''''''''''''''
'  Function to search for occurences of a sub-string.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function SubStrCount(sStr As String, _
                            sSubStr As String, _
                            Optional lStartPos As Long = 1, _
                            Optional ByVal lCompare As _
                            VbCompareMethod = vbBinaryCompare, _
                            Optional ByVal bWordOnly As Boolean, _
                            Optional ByVal lHitLimit As Long _
                            ) As Long ' ©Rd

    Dim sStrV As String, sSubStrV As String
    Dim lLenStr As Long, lLenSub As Long
    Dim lBefore As Long, lAfter As Long
    Dim lStartV As Long, lHit As Long
    Dim lOffStart As Long, bDelim As Boolean

    On Error GoTo FreakOut

    lLenStr = Len(sStr)
    If (lLenStr = 0) Then Exit Function ' No text

    lLenSub = Len(sSubStr)
    If (lLenSub = 0) Then Exit Function ' Nothing to find

    If (lStartPos < 1) Then lHit = 1 Else lHit = lStartPos

    If (lCompare = vbTextCompare) Then
        ' Better to convert once to lowercase than on every call to InStr
        sSubStrV = LCase$(sSubStr): sStrV = LCase$(sStr)

        lHit = InStr(lHit, sStrV, sSubStrV, vbBinaryCompare)
    Else                         ' Do first search
        lHit = InStr(lHit, sStr, sSubStr, vbBinaryCompare)
    End If

    Do While (lHit)    ' Do until no more hits

        If bWordOnly = False Then

            lStartV = lHit
            SubStrCount = SubStrCount + 1
            If (SubStrCount = lHitLimit) Then Exit Do

            lOffStart = lLenSub ' Offset next start pos
        Else
            lOffStart = 1 ' Default offset start pos

            lBefore = lHit - 1
            If (lBefore = 0) Then
                bDelim = True
            Else
                bDelim = IsDelim(Mid$(sStr, lBefore, 1))
            End If

            If bDelim Then

                lAfter = lHit + lLenSub
                If (lAfter > lLenStr) Then
                    bDelim = True
                Else
                    bDelim = IsDelim(Mid$(sStr, lAfter, 1))
                End If

                If bDelim Then

                    lStartV = lHit
                    SubStrCount = SubStrCount + 1
                    If (SubStrCount = lHitLimit) Then Exit Do

                    lOffStart = lLenSub ' Offset next start pos
                End If
            End If
        End If

        If (lCompare = vbTextCompare) Then
            lHit = InStr(lHit + lOffStart, sStrV, sSubStrV)
        Else
            lHit = InStr(lHit + lOffStart, sStr, sSubStr)
        End If
    Loop

    If (lHitLimit > 0) Then lStartPos = lStartV
FreakOut:
End Function

'IsDelim'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'  This function is called by the SubStrCount function during a
'  whole word only procedure. You can also call this function
'  from your own code - very handy for string parsing.
'
'  It checks if the character passed is a common word delimiter,
'  and then returns True or False accordingly.
'
'  By default, any non-alphabetic character is considered a word
'  delimiter, including underscores, apostrophes and numbers.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsDelim(Char As String) As Boolean ' ©Rd
    Select Case Asc(Char)
        ' Uppercase, Underscore, Lowercase chars not delimiters
        Case 65 To 90, 95, 97 To 122: IsDelimI = False

        'Case 39, 146: IsDelim = False  ' Apostrophes not delimiters
        'Case 48 To 57: IsDelim = False ' Numeric chars not delimiters

        Case Else: IsDelim = True ' Any other character is delimiter
    End Select
    If (IsDelim) And Not (LenB(mNotDelim) = 0) Then
        If Not (InStr(mNotDelim, Char) = 0) Then
            IsDelim = False
            ' SetDelim doesn't allow chars to repeat
            ' in both lists so we can exit
            Exit Function
        End If
    End If
    If Not (IsDelim) And Not (LenB(mDelimList) = 0) Then
        ' May need alphabetic characters to behave as delimiters
        IsDelim = Not (InStr(mDelimList, Char) = 0)
    End If
End Function

'SetDelim''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'  Specifies whether character(s) should be handled as delimiter
'  in whole word only searches.
'
'  Note that all non-alphabetic characters are already treated
'  as word delimiters by default and do not need to be specified
'  through SetDelim.
'
'  Note that multiple characters must not be seperated by spaces
'  or any other character.
'
'  For example, to set all numeric charaters, underscores and
'  apostrophes as part of whole words (non-delimiters):
'
'  SetDelim "1234567890_'’", False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetDelim(Char As String, Optional IsDelimiter As Boolean) ' ©Rd
    On Error GoTo ErrHandler
    Dim iPos As Long, sChr As String
    Dim fIs As Boolean, fNot As Boolean
    Dim idx1 As Long, idx2 As Long

    fIs = Not (LenB(mDelimList) = 0)
    fNot = Not (LenB(mNotDelim) = 0)

    For iPos = 1 To Len(Char)
        sChr = Mid$(Char, iPos, 1)
        If fIs Then idx1 = InStr(mDelimList, sChr)
        If fNot Then idx2 = InStr(mNotDelim, sChr)

        If IsDelimiter Then
            If (idx1 = 0) Then mDelimList = mDelimList & sChr
            If Not (idx2 = 0) Then
                mNotDelim = Left$(mNotDelim, idx2 - 1) & Mid$(mNotDelim, idx2 + 1)
            End If
        Else
            If (idx2 = 0) Then mNotDelim = mNotDelim & sChr
            If Not (idx1 = 0) Then
                mDelimList = Left$(mDelimList, idx1 - 1) & Mid$(mDelimList, idx1 + 1)
            End If
        End If
    Next iPos
ErrHandler:
End Sub

' Rd - crYptic but cRaZy!                                      :)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
