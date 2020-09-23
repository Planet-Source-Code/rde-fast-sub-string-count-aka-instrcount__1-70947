Attribute VB_Name = "mInStrCnt"
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal lByteLen As Long)

Public Static Function InStrCount04( _
    ByRef Text As String, _
    ByRef Find As String, _
    Optional ByVal Start As Long = 1, _
    Optional ByVal Compare As VbCompareMethod = vbBinaryCompare _
  ) As Long
' by Jost Schwider, jost@schwider.de, 20010912, rev 001 20011121
  Const MODEMARGIN = 8
  Dim TextAsc() As Integer
  Dim TextData As Long
  Dim TextPtr As Long
  Dim FindAsc(0 To MODEMARGIN) As Integer
  Dim FindLen As Long
  Dim FindChar1 As Integer
  Dim FindChar2 As Integer
  Dim i As Long

  If Compare = vbBinaryCompare Then
    FindLen = Len(Find)
    If FindLen Then
      'Ersten Treffer bestimmen:
      If Start < 2 Then
        Start = InStrB(Text, Find)
      Else
        Start = InStrB(Start + Start - 1, Text, Find)
      End If

      If Start Then
        InStrCount04 = 1
        If FindLen <= MODEMARGIN Then

          If TextPtr = 0 Then
            'TextAsc-Array vorbereiten:
            ReDim TextAsc(1 To 1)
            TextData = VarPtr(TextAsc(1))
            CopyMemory TextPtr, ByVal ArrayPtr(TextAsc), 4
            TextPtr = TextPtr + 12
          End If

          'TextAsc-Array initialisieren:
          CopyMemory ByVal TextPtr, ByVal VarPtr(Text), 4 'pvData
          CopyMemory ByVal TextPtr + 4, Len(Text), 4      'nElements

          Select Case FindLen
          Case 1

            'Das Zeichen buffern:
            FindChar1 = AscW(Find)

            'Zählen:
            For Start = Start \ 2 + 2 To Len(Text)
              If TextAsc(Start) = FindChar1 Then InStrCount04 = InStrCount04 + 1
            Next Start

          Case 2

            'Beide Zeichen buffern:
            FindChar1 = AscW(Find)
            FindChar2 = AscW(Right$(Find, 1))

            'Zählen:
            For Start = Start \ 2 + 3 To Len(Text) - 1
              If TextAsc(Start) = FindChar1 Then
                If TextAsc(Start + 1) = FindChar2 Then
                  InStrCount04 = InStrCount04 + 1
                  Start = Start + 1
                End If
              End If
            Next Start

          Case Else

            'FindAsc-Array füllen:
            CopyMemory ByVal VarPtr(FindAsc(0)), ByVal StrPtr(Find), FindLen + FindLen
            FindLen = FindLen - 1

            'Die ersten beiden Zeichen buffern:
            FindChar1 = FindAsc(0)
            FindChar2 = FindAsc(1)

            'Zählen:
            For Start = Start \ 2 + 2 + FindLen To Len(Text) - FindLen
              If TextAsc(Start) = FindChar1 Then
                If TextAsc(Start + 1) = FindChar2 Then
                  For i = 2 To FindLen
                    If TextAsc(Start + i) <> FindAsc(i) Then Exit For
                  Next i
                  If i > FindLen Then
                    InStrCount04 = InStrCount04 + 1
                    Start = Start + FindLen
                  End If
                End If
              End If
            Next Start

          End Select

          'TextAsc-Array restaurieren:
          CopyMemory ByVal TextPtr, TextData, 4 'pvData
          CopyMemory ByVal TextPtr + 4, 1&, 4   'nElements

        Else

          'Konventionell Zählen:
          FindLen = FindLen + FindLen
          Start = InStrB(Start + FindLen, Text, Find)
          Do While Start
            InStrCount04 = InStrCount04 + 1
            Start = InStrB(Start + FindLen, Text, Find)
          Loop

        End If 'FindLen <= MODEMARGIN
      End If 'Start
    End If 'FindLen
  Else
    'Groß-/Kleinschreibung ignorieren:
    InStrCount04 = InStrCount04(LCase$(Text), LCase$(Find), Start)
  End If
End Function

Public Function InStrCnt(sSrc As String, sTerm As String, Optional ByVal lStart As Long = 1, _
                                             Optional CaseSensitive As Boolean = True) As Long
  'By Jost Schwider, jost@schwider.de, 20010912, rev 001 20011121
  Const MODE_MARGIN As Long = 8
  Dim vbMode As VbCompareMethod
  vbMode = IIf(CaseSensitive, vbBinaryCompare, vbTextCompare)
  Dim ipSrcAsc() As Integer
  Dim lpSrcData As Long
  Dim lpSrcArrPtr As Long
  Dim lLenTerm As Long
  Dim ipTermChr1 As Integer
  Dim ipTermChr2 As Integer

  If vbMode = vbBinaryCompare Then
    lLenTerm = Len(sTerm)
    If lLenTerm Then

      'Ersten Treffer bestimmen:
      lStart = InStrB(lStart + lStart - 1, sSrc, sTerm) ' Search for term (binary/byte level)

      If lStart Then
        InStrCnt = 1
        If lLenTerm <= MODE_MARGIN Then

          If lpSrcData = 0 Then
            'TextAsc-Array vorbereiten:
            ReDim ipSrcAsc(1 To 1)
            lpSrcData = VarPtr(ipSrcAsc(1))
            CopyMemory lpSrcArrPtr, ByVal ArrayPtr(ipSrcAsc), 4
            lpSrcArrPtr = lpSrcArrPtr + 12 ' First 12 array info
          End If

          'TextAsc-Array initialisieren:
          CopyMemory ByVal lpSrcArrPtr, ByVal VarPtr(sSrc), 4 'pvStrData
          CopyMemory ByVal lpSrcArrPtr + 4, Len(sSrc), 4      'nElements

          Select Case lLenTerm
          Case 1

            'Das Zeichen buffern:
            ipTermChr1 = AscW(sTerm)

            'Zählen:
            For lStart = lStart \ 2 + 2 To Len(sSrc)
              If ipSrcAsc(lStart) = ipTermChr1 Then InStrCnt = InStrCnt + 1
            Next lStart

          Case 2

            'Beide Zeichen buffern:
            ipTermChr1 = AscW(sTerm)
            ipTermChr2 = AscW(Right$(sTerm, 1))

            'Zählen:
            For lStart = lStart \ 2 + 3 To Len(sSrc) - 1
              If ipSrcAsc(lStart) = ipTermChr1 Then
                If ipSrcAsc(lStart + 1) = ipTermChr2 Then
                  InStrCnt = InStrCnt + 1
                  lStart = lStart + 1
                End If
              End If
            Next lStart

          Case Else

            'iaTermAsc-Array füllen:
            Dim iaTermAsc(0 To MODE_MARGIN) As Integer
            CopyMemory ByVal VarPtr(iaTermAsc(0)), ByVal StrPtr(sTerm), lLenTerm + lLenTerm
            lLenTerm = lLenTerm - 1

            'Die ersten beiden Zeichen buffern:
            ipTermChr1 = iaTermAsc(0)
            ipTermChr2 = iaTermAsc(1)

            'Zählen:
            Dim i As Long
            For lStart = lStart \ 2 + 2 + lLenTerm To Len(sSrc) - lLenTerm
              If ipSrcAsc(lStart) = ipTermChr1 Then
                If ipSrcAsc(lStart + 1) = ipTermChr2 Then
                  For i = 2 To lLenTerm
                    If ipSrcAsc(lStart + i) <> iaTermAsc(i) Then Exit For
                  Next i
                  If i > lLenTerm Then
                    InStrCnt = InStrCnt + 1
                    lStart = lStart + lLenTerm
                  End If
                End If
              End If
            Next lStart

          End Select

          'ipSrcAsc-Array restaurieren:
          CopyMemory ByVal lpSrcArrPtr, lpSrcData, 4 'pvData
          CopyMemory ByVal lpSrcArrPtr + 4, 1&, 4   'nElements

        Else

          'Konventionell Zählen:
          lLenTerm = lLenTerm + lLenTerm
          lStart = InStrB(lStart + lLenTerm, sSrc, sTerm)
          Do While lStart
            InStrCnt = InStrCnt + 1
            lStart = InStrB(lStart + lLenTerm, sSrc, sTerm)
          Loop

        End If 'lLenTerm <= iscMODE_MARGIN (8 bytes)
      End If 'lStart
    End If 'lLenTerm
  Else
    'Groß-/Kleinschreibung ignorieren:
    InStrCnt = InStrCnt(LCase$(sSrc), LCase$(sTerm), lStart)
  End If
End Function


' + ArrayPtr ++++++++++++++++++++++++++++++++++++++++++++

' This function returns a pointer to the SAFEARRAY header of
' any Visual Basic array, including a Visual Basic string array.

' Substitutes both ArrPtr and StrArrPtr.

' This function will work with vb5 or vb6 without modification.

Public Function ArrayPtr(Arr) As Long
    Dim iDataType As Integer
    On Error GoTo UnInit
    CopyMemory iDataType, Arr, 2&                       ' get the real VarType of the argument, this is similar to VarType(), but returns also the VT_BYREF bit
    If (iDataType And vbArray) = vbArray Then           ' if a valid array was passed
        CopyMemory ArrayPtr, ByVal VarPtr(Arr) + 8&, 4& ' get the address of the SAFEARRAY descriptor stored in the second half of the Variant parameter that has received the array. Thanks to Francesco Balena.
    End If
UnInit:
End Function

' +++++++++++++++++++++++++++++++++++++++++++++++++++++++
