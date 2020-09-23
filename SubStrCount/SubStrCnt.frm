VERSION 5.00
Begin VB.Form frmSubStrCnt 
   Caption         =   "Sub-String Count Ex"
   ClientHeight    =   5220
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4335
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5220
   ScaleWidth      =   4335
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdTest 
      Caption         =   "Performance Test..."
      Height          =   345
      Left            =   180
      TabIndex        =   12
      Top             =   3300
      Width           =   1725
   End
   Begin VB.TextBox txtStart 
      Height          =   285
      Left            =   1140
      TabIndex        =   10
      Top             =   2850
      Width           =   555
   End
   Begin VB.TextBox txtLimit 
      Height          =   285
      Left            =   3330
      TabIndex        =   8
      Top             =   2850
      Width           =   465
   End
   Begin VB.TextBox txtHits 
      Height          =   255
      Left            =   3300
      TabIndex        =   6
      Top             =   2370
      Width           =   465
   End
   Begin VB.CheckBox chkWWonly 
      Caption         =   "Whole Words Only"
      Height          =   255
      Left            =   420
      TabIndex        =   4
      Top             =   2520
      Width           =   2055
   End
   Begin VB.CheckBox chkCaseSens 
      Caption         =   "Case Sensitive"
      Height          =   195
      Left            =   570
      TabIndex        =   3
      Top             =   2220
      Width           =   1695
   End
   Begin VB.CommandButton cmdCount 
      Caption         =   "Count Now..."
      Height          =   345
      Left            =   2850
      TabIndex        =   2
      Top             =   1830
      Width           =   1155
   End
   Begin VB.TextBox txtSubText 
      Height          =   315
      Left            =   1710
      TabIndex        =   1
      Top             =   1830
      Width           =   1005
   End
   Begin VB.TextBox txtSearchText 
      Height          =   1395
      HideSelection   =   0   'False
      Left            =   330
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   210
      Width           =   3675
   End
   Begin VB.Label Label5 
      Caption         =   "mInStrCnt.InStrCount04:"
      Height          =   210
      Index           =   3
      Left            =   210
      TabIndex        =   22
      Top             =   4290
      Width           =   2205
   End
   Begin VB.Label Label5 
      Height          =   210
      Index           =   8
      Left            =   2520
      TabIndex        =   21
      Top             =   4290
      Width           =   2205
   End
   Begin VB.Label Label5 
      Height          =   210
      Index           =   10
      Left            =   2520
      TabIndex        =   20
      Top             =   4830
      Width           =   1965
   End
   Begin VB.Label Label5 
      Height          =   210
      Index           =   9
      Left            =   2520
      TabIndex        =   19
      Top             =   4560
      Width           =   2025
   End
   Begin VB.Label Label5 
      Height          =   210
      Index           =   7
      Left            =   2520
      TabIndex        =   18
      Top             =   4020
      Width           =   2205
   End
   Begin VB.Label Label5 
      Height          =   210
      Index           =   6
      Left            =   2520
      TabIndex        =   17
      Top             =   3750
      Width           =   2205
   End
   Begin VB.Label Label5 
      Caption         =   "Your InStrCnt:"
      Height          =   210
      Index           =   5
      Left            =   720
      TabIndex        =   16
      Top             =   4830
      Width           =   1695
   End
   Begin VB.Label Label5 
      Caption         =   "mInStrCnt.InStrCnt:"
      Height          =   210
      Index           =   4
      Left            =   480
      TabIndex        =   15
      Top             =   4560
      Width           =   2025
   End
   Begin VB.Label Label5 
      Caption         =   "mSubStrCnt.SubStringCount:"
      Height          =   210
      Index           =   2
      Left            =   180
      TabIndex        =   14
      Top             =   4020
      Width           =   2205
   End
   Begin VB.Label Label5 
      Caption         =   "mSubStrCnt.SubStrCount:"
      Height          =   210
      Index           =   1
      Left            =   390
      TabIndex        =   13
      Top             =   3750
      Width           =   2205
   End
   Begin VB.Label Label4 
      Caption         =   "Start At:"
      Height          =   285
      Left            =   420
      TabIndex        =   11
      Top             =   2880
      Width           =   1065
   End
   Begin VB.Label Label3 
      Caption         =   "Limit Hit Count To:"
      Height          =   225
      Left            =   1890
      TabIndex        =   9
      Top             =   2880
      Width           =   2085
   End
   Begin VB.Label Label2 
      Caption         =   "Hits:"
      Height          =   195
      Left            =   2850
      TabIndex        =   7
      Top             =   2400
      Width           =   765
   End
   Begin VB.Label Label1 
      Caption         =   "Search Sub-String:"
      Height          =   225
      Left            =   300
      TabIndex        =   5
      Top             =   1860
      Width           =   2115
   End
End
Attribute VB_Name = "frmSubStrCnt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function PerfCount Lib "kernel32" Alias "QueryPerformanceCounter" (lpPerformanceCount As Currency) As Long
Private Declare Function PerfFreq Lib "kernel32" Alias "QueryPerformanceFrequency" (lpFrequency As Currency) As Long
Private mCurFreq As Currency

Private lStart As Long

Private Function ProfileStart() As Currency
    If mCurFreq = 0 Then PerfFreq mCurFreq
    If (mCurFreq) Then PerfCount ProfileStart
End Function

Private Function ProfileStop(ByVal curStart As Currency) As Currency
    If (mCurFreq) Then
        Dim curStop As Currency
        PerfCount curStop
        ProfileStop = (curStop - curStart) / mCurFreq ' cpu tick accurate
        curStop = 0
    End If
End Function

Private Function OpenFile(sFileSpec As String) As String
    ' Handle errors if they occur
    On Error GoTo GetFileError
    Dim iFile As Integer
    iFile = FreeFile
    ' Open in binary mode, let others read but not write
    Open sFileSpec For Binary Access Read Lock Write As #iFile
    ' Allocate the length first
    OpenFile = Space$(LOF(iFile))
    ' Get the file in one chunk
    Get #iFile, , OpenFile
GetFileError:
    Close #iFile ' Close the file
End Function

Private Sub cmdTest_Click()
    Dim curElapse As Currency
    Dim sFile As String
    Dim r1 As Single
    Dim i As Long

    For i = 6 To 10
        Label5(i) = vbNullString
        Label5(i).Refresh
    Next i

    Screen.MousePointer = vbHourglass
    sFile = OpenFile(App.Path & "\SubStrCnt.bas")

    curElapse = ProfileStart
    For i = 1 To 10000
        txtHits = SubStrCount(sFile, "is", 1, Abs(chkCaseSens - 1), CBool(chkWWonly))
    Next i
    r1 = CSng(ProfileStop(curElapse))
    Label5(6) = Format$(r1, "##0.0000") & " secs"
    Label5(6).Refresh
    
    curElapse = ProfileStart
    For i = 1 To 10000
        txtHits = SubStringCount(sFile, "is", 1, Abs(chkCaseSens - 1), CBool(chkWWonly))
    Next i
    r1 = CSng(ProfileStop(curElapse))
    Label5(7) = Format$(r1, "##0.0000") & " secs"
    Label5(7).Refresh

    If chkWWonly = 0 Then

        curElapse = ProfileStart
        For i = 1 To 10000
            txtHits = InStrCount04(sFile, "is", 1, Abs(chkCaseSens - 1))
        Next i
        r1 = CSng(ProfileStop(curElapse))
        Label5(8) = Format$(r1, "##0.0000") & " secs"
        Label5(8).Refresh

        curElapse = ProfileStart
        For i = 1 To 10000
            txtHits = InStrCnt(sFile, "is", 1, CBool(chkCaseSens))
        Next i
        r1 = CSng(ProfileStop(curElapse))
        Label5(9) = Format$(r1, "##0.0000") & " secs"
        Label5(9).Refresh

    Else
        Label5(8) = "999999"
        Label5(8).Refresh
        Label5(9) = "999999"
        Label5(9).Refresh
    End If
    
'    curElapse = ProfileStart
'    For i = 1 To 10000 'YourInStrCnt
'        txtHits = YourInStrCnt(sFile, "is", 1, Abs(chkCaseSens - 1), CBool(chkWWonly))
'    Next i
'    r1 = CSng(ProfileStop(curElapse))
'    Label5(10) = Format$(r1, "##0.0000") & " secs"
'    Label5(10).Refresh

    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
  txtSearchText = "This is the Search Text2search for occurences of text sub-strings that may be searched in this search text. This search Text is only 278Text characters in the txtSearchText text Control. I hope this Search_text will help you with your Text context searching research."
  txtSubText = "Text"
  txtStart = "1"
  txtLimit = "1"
End Sub

Private Sub cmdCount_Click()
'count now
  txtSearchText.SelLength = 0
  lStart = txtStart
  'txtHits = SubStrCount(txtSearchText, txtSubText, lStart, Abs(chkCaseSens - 1), CBool(chkWWonly), txtLimit)
  txtHits = SubStringCount(txtSearchText, txtSubText, lStart, Abs(chkCaseSens - 1), CBool(chkWWonly), txtLimit)

  If txtLimit > 0 And lStart > 0 Then
      txtSearchText.SelStart = lStart - 1 ' zero based sel
      txtSearchText.SelLength = Len(txtSubText)
      txtStart = lStart + 1
  Else
      txtStart = "0"
  End If
End Sub

Private Sub txtSearchText_Click()
    txtStart = txtSearchText.SelStart + 1 ' zero based sel
End Sub

Private Sub txtSearchText_KeyUp(KeyCode As Integer, Shift As Integer)
    If (txtSearchText.SelLength > 0) Then
        txtSubText = Trim$(txtSearchText.SelText)
    End If
End Sub

Private Sub txtSearchText_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (txtSearchText.SelLength > 0) Then
        txtSubText = Trim$(txtSearchText.SelText)
    End If
End Sub
