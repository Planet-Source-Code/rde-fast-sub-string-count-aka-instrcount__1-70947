Attribute VB_Name = "mSubStrCntApi"
Option Explicit

' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤-©Rd-¤
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
'  SubStringCount will return the number of matches or zero if none.
'
'                     Extended Functionality
'
'  This SubStringCount implementation offers extended functionality
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
'  make subsequent calls to SubStringCount by passing lStartPos + 1
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
' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤

' Declare some CopyMemory Alias's (thanks Bruce :)
Private Declare Sub CopyMemByV Lib "kernel32" Alias "RtlMoveMemory" (ByVal lpDest As Long, ByVal lpSrc As Long, ByVal lByteLen As Long)
Private Declare Sub CopyMemByR Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal lByteLen As Long)

Private lDelimList As Long
Private lNotDelim As Long

Private laDelim() As Long
Private laNotDel() As Long

' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
'       Function to search for occurences of a sub-string.
' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤

Public Function SubStringCount(sStr As String, _
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
    Dim lStart As Long, lHit As Long
    Dim lOffStart As Long, bDelim As Boolean

    On Error GoTo FreakOut

    lLenStr = Len(sStr)
    If (lLenStr = 0) Then Exit Function ' No text

    lLenSub = Len(sSubStr)
    If (lLenSub = 0) Then Exit Function ' Nothing to find

    If lStartPos < 1 Then lHit = 1 Else lHit = lStartPos ' Validate start pos

    If (lCompare = vbTextCompare) Then
        ' Better to convert once to lowercase than on every call to InStr
        sSubStrV = LCase$(sSubStr): sStrV = LCase$(sStr)
    Else
        CopyMemByV VarPtr(sSubStrV), VarPtr(sSubStr), 4&
        CopyMemByV VarPtr(sStrV), VarPtr(sStr), 4&
    End If

    lHit = InStr(lHit, sStrV, sSubStrV, vbBinaryCompare)

    Do While (lHit)           ' Do until no more hits

        If bWordOnly = False Then

            lStart = lHit
            SubStringCount = SubStringCount + 1
            If (SubStringCount = lHitLimit) Then Exit Do

            lOffStart = lLenSub ' Offset next start pos
        Else
            lOffStart = 1 ' Default offset start pos

            lBefore = lHit - 1
            If (lBefore = 0) Then
                bDelim = True
            Else
                bDelim = IsDelimI(MidI(sStrV, lBefore))
            End If

            If bDelim Then

                lAfter = lHit + lLenSub
                If (lAfter > lLenStr) Then
                    bDelim = True
                Else
                    bDelim = IsDelimI(MidI(sStrV, lAfter))
                End If

                If bDelim Then

                    lStart = lHit
                    SubStringCount = SubStringCount + 1
                    If (SubStringCount = lHitLimit) Then Exit Do

                    lOffStart = lLenSub ' Offset next start pos
                End If
            End If
        End If

        lHit = InStr(lHit + lOffStart, sStrV, sSubStrV)
    Loop

    If (lHitLimit > 0) Then lStartPos = lStart
FreakOut:
    If (lCompare = vbBinaryCompare) Then
        CopyMemByR ByVal VarPtr(sSubStrV), 0&, 4& ' De-reference pointer
        CopyMemByR ByVal VarPtr(sStrV), 0&, 4&    ' De-reference pointer
    End If
End Function

' ¤¤ IsDelim ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
'
'  This function is called by the Replace function during a
'  whole word only procedure. You can also call this function
'  from your own code - very handy for string parsing.
'
'  It checks if the character passed is a common word delimiter,
'  and then returns True or False accordingly.
'
'  By default, any non-alphabetic character (except for an
'  underscore) is considered a word delimiter, including numbers.
'
'  By default, an underscore is treated as part of a whole word,
'  and so is not considered a word delimiter in whole word only
'  searches.
'
' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤

Public Function IsDelim(Char As String) As Boolean ' ©Rd
    Dim iIdx As Long
    Dim iAscW As Long
    iAscW = AscW(Char)
    Select Case iAscW
        ' Uppercase, Underscore, Lowercase chars not delimiters
        Case 65 To 90, 95, 97 To 122: IsDelim = False

        'Case 39, 146: IsDelim = False  ' Apostrophes not delimiters
        'Case 48 To 57: IsDelim = False ' Numeric chars not delimiters

        Case Else: IsDelim = True ' Any other character is delimiter
    End Select
    If IsDelim And (lNotDelim <> 0) Then
        Do Until iIdx = lNotDelim
            If laNotDel(iIdx) = iAscW Then Exit Do
            iIdx = iIdx + 1
        Loop
        If (iIdx < lNotDelim) Then
            IsDelim = False
            ' SetDelim doesn't allow chars to repeat
            ' in both lists so we can exit
            Exit Function
        End If
    End If
    If (IsDelim = False) And (lDelimList <> 0) Then
        ' May need alphabetic characters to behave as delimiters
        Do Until iIdx = lDelimList
            If laDelim(iIdx) = iAscW Then Exit Do
            iIdx = iIdx + 1
        Loop
        IsDelim = iIdx < lDelimList
    End If
End Function

' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤

Public Function IsDelimI(ByVal iAscW As Long) As Boolean ' ©Rd
    Dim iIdx As Long
    Select Case iAscW
        ' Uppercase, Underscore, Lowercase chars not delimiters
        Case 65 To 90, 95, 97 To 122: IsDelimI = False

        'Case 39, 146: IsDelim = False  ' Apostrophes not delimiters
        'Case 48 To 57: IsDelim = False ' Numeric chars not delimiters

        Case Else: IsDelimI = True ' Any other character is delimiter
    End Select
    If IsDelimI And (lNotDelim <> 0) Then
        Do Until iIdx = lNotDelim
            If laNotDel(iIdx) = iAscW Then Exit Do
            iIdx = iIdx + 1
        Loop
        If (iIdx < lNotDelim) Then
            IsDelimI = False
            ' SetDelim doesn't allow chars to repeat
            ' in both lists so we can exit
            Exit Function
        End If
    End If
    If (IsDelimI = False) And (lDelimList <> 0) Then
        ' May need alphabetic characters to behave as delimiters
        iIdx = 0
        Do Until iIdx = lDelimList
            If laDelim(iIdx) = iAscW Then Exit Do
            iIdx = iIdx + 1
        Loop
        IsDelimI = iIdx < lDelimList
    End If
End Function

' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤

Public Property Get MidI(sStr As String, ByVal lPos As Long) As Integer
    CopyMemByR MidI, ByVal StrPtr(sStr) + lPos + lPos - 2, 2&
End Property

Public Property Get MidIB(sStr As String, ByVal lPosB As Long) As Integer
    CopyMemByR MidIB, ByVal StrPtr(sStr) + lPosB - 1, 2&
End Property

' ¤¤ SetDelim ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
'
'  Specifies whether character(s) should be handled as delimiter
'  in whole word only searches.
'
'  Remember, all non-alphabetic characters (with the exception of
'  underscores) are already treated as word delimiters by default
'  and do not need to be specified through SetDelim.
'
'  Note that multiple characters must not be seperated by spaces
'  or any other character.
'
'  For example, to set all numeric charaters and apostrophes as
'  part of whole words (non-delimiters):
'
'  SetDelim "1234567890'’", False
'
' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤

Public Sub SetDelim(Char As String, Optional ByVal IsDelimiter As Boolean) ' ©Rd
    On Error GoTo ErrHandler
    Dim iPos As Long, iChr As Long
    Dim idx1 As Long, idx2 As Long

    idx1 = Len(Char)
    If IsDelimiter Then
        ReDim Preserve laDelim(0 To lDelimList + idx1) As Long
    Else
        ReDim Preserve laNotDel(0 To lNotDelim + idx1) As Long
    End If
    For iPos = 1 To idx1
        iChr = MidI(Char, iPos)
        idx1 = 0
        Do Until idx1 = lDelimList
            If laDelim(idx1) = iChr Then Exit Do
            idx1 = idx1 + 1
        Loop
        idx2 = 0
        Do Until idx2 = lNotDelim
            If laNotDel(idx2) = iChr Then Exit Do
            idx2 = idx2 + 1
        Loop
        If IsDelimiter Then
            If (idx1 = lDelimList) Then
                laDelim(lDelimList) = iChr
                lDelimList = lDelimList + 1
            End If
            If (idx2 < lNotDelim) Then
                lNotDelim = lNotDelim - 1
                laNotDel(idx2) = laNotDel(lNotDelim)
            End If
        Else
            If (idx2 = lNotDelim) Then
                laNotDel(lNotDelim) = iChr
                lNotDelim = lNotDelim + 1
            End If
            If (idx1 < lDelimList) Then
                lDelimList = lDelimList - 1
                laDelim(idx1) = laDelim(lDelimList)
            End If
        End If
    Next iPos
ErrHandler:
End Sub

'  Rd - crYptic but cRaZy                                      :)
' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
'
