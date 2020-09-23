VERSION 5.00
Begin VB.MDIForm mdiScreen 
   BackColor       =   &H8000000C&
   Caption         =   "Tape Backup Manager (Required Backup Software from Microsoft Windows NT and a Tape Recorder)"
   ClientHeight    =   9465
   ClientLeft      =   255
   ClientTop       =   750
   ClientWidth     =   13755
   Icon            =   "mdiScreen.frx":0000
   LinkTopic       =   "MDIForm1"
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2280
      Top             =   960
   End
   Begin VB.Menu opSchedule 
      Caption         =   "Schedule..."
   End
   Begin VB.Menu opException 
      Caption         =   "Exception..."
   End
   Begin VB.Menu mnuFiles 
      Caption         =   "Files..."
   End
   Begin VB.Menu mnuTapeInfo 
      Caption         =   "TapeInfo..."
   End
   Begin VB.Menu mnuServices 
      Caption         =   "Services"
      Visible         =   0   'False
      Begin VB.Menu mnuAddRem 
         Caption         =   "Add to services"
      End
   End
   Begin VB.Menu mnuBackup 
      Caption         =   "Backup Now"
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "About..."
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Popups"
      Visible         =   0   'False
      Begin VB.Menu mnuShow 
         Caption         =   "Show"
      End
      Begin VB.Menu mnuseparator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "mdiScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents gSysTray As clsSysTray
Attribute gSysTray.VB_VarHelpID = -1
Private Sub gSysTray_LButtonDblClk()
    On Error Resume Next
    Me.WindowState = vbNormal
    Me.Show
End Sub

Private Sub gSysTray_RButtonUP()
    PopupMenu Me.mnuPopup
End Sub
Private Sub MDIForm_Load()
    Set gSysTray = New clsSysTray
    gSysTray.LoadIcon Me.Icon, Me
    gSysTray.ToolTip = "TapeManager" & Chr(0)
    gSysTray.IconInSysTray
End Sub

Private Sub MDIForm_Resize()
    If Me.WindowState = vbMinimized Then Me.Hide
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    Cancel = 1
    Me.WindowState = vbMinimized
    Me.Hide
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show
End Sub



Private Sub mnuBackup_Click()
    runBackup
End Sub

Private Sub mnuExit_Click()
    End
End Sub

Private Sub mnuFiles_Click()
    frmJobManager.Show
End Sub

Private Sub mnuShow_Click()
    On Error Resume Next
    Me.WindowState = vbNormal
    Me.Show
End Sub

Private Sub mnuTapeInfo_Click()
    frmStatus.Show
End Sub

Private Sub opException_Click()
    frmException.Show
End Sub

Private Sub opSchedule_Click()
    frmScheduler.Show
End Sub


Private Sub sdf_Click()

End Sub

Private Sub Timer1_Timer()
    If (checkSchedule()) Then
        runBackup
    End If
End Sub
Private Sub runBackup()
On Error Resume Next
    Dim strTemp As String
    Dim strFiles As String
    Open App.Path & "\files.bak" For Input As #1
    Do
        strTemp = ""
        Line Input #1, strTemp
        If (strTemp <> "") Then
            strFiles = strFiles & " """ & strTemp & """"
        End If
    Loop While (strTemp <> "")
    Close #1
    If (strFiles <> "") Then
        Open App.Path & "\jobs.bak" For Input As #1
        Line Input #1, strTemp
        If (strTemp <> "") Then
            Shell "ntbackup backup" & strFiles & " /j """ & strTemp & """ /p ""DLT"" /m copy /um /l:f /hc:on"
        Else
            Shell "ntbackup backup" & strFiles & " /j ""UNTITLED"" /p ""DLT"" /m copy /um /l:f /hc:on"
        End If
    End If
End Sub
Private Function checkSchedule() As Boolean
    On Error Resume Next
    Dim strTemp As String
    Dim strSplit() As String
    Dim strBuffer As String
    Dim lngSepare As Long
    Dim lngBuffer As Long
    Dim strValues() As String
    Open App.Path & "\schedule.bak" For Input As #1
    Do
        strTemp = ""
        Line Input #1, strTemp
        If (strTemp <> "") Then
            strSplit = Split(strTemp, "|")
            If (strSplit(1) <> "") Then
                Dim strSepare() As String
                strSepare = Split(strSplit(1), "@")
                lngSepare = DateTime.DateDiff("d", CDate(strSepare(1)), DateTime.Now)
                If (lngSepare Mod CLng(strSepare(0)) = 0 And lngSepare > 0) Then
                    If (strSplit(0) = DateTime.Time) Then
                        If Not checkException Then
                            Close #1
                            checkSchedule = True
                        End If
                    End If
                End If
            End If
            If (strSplit(2) <> "") Then
                strValues = Split(strSplit(2), ":")
                If (DateTime.DateDiff("d", CDate(strValues(1)), DateTime.Date) Mod (7 * CLng(strValues(0))) = 0) Then
                    If (strSplit(0) = DateTime.Time) Then
                        If Not checkException Then
                            Close #1
                            checkSchedule = True
                        End If
                    End If
                End If
            End If
            If (strSplit(3) <> "") Then
                strValues = Split(strSplit(3), ":")
                If (strValues(0) = "d") Then
                    If (DateTime.DateSerial(DateTime.Year(DateTime.Now), DateTime.Month(DateTime.Now), strValues(1)) = DateTime.Date) Then
                        If (strSplit(0) = DateTime.Time) Then
                            If Not checkException Then
                                Close #1
                                checkSchedule = True
                            End If
                        End If
                    End If
                Else
'                    Dim dtDate As Date
'                    Dim intWeekDay As Integer
'                    Select Case strValues(1)
'                        Case "1"
'
'                        Case "2"
'
'                        Case "3"
'
'                        Case "4"
'
'                        Case "5"
'
'                    End Select
'                    dtDate = DateTime.DateSerial(DateTime.Year(DateTime.Now), DateTime.Month(DateTime.Now), 1)
'                    intWeekDay = DateTime.Weekday(dtDate)
'                    If (intWeekDay > CInt(strValues(2))) Then
'                        dtDate = DateTime.DateAdd("d", 7 - (intWeekDay - CInt(strValues(2))), dtDate)
'                    Else
'                        dtDate = DateTime.DateAdd("d", CInt(strValues(2)) - intWeekDay, dtDate)
'                    End If
'                    If (dtDate = DateTime.Date) Then
'                        If (strSplit(0) = DateTime.Time) Then
'                            If Not checkException Then
'                                Close #1
'                                checkSchedule = True
'                            End If
'                        End If
'                    End If
                End If
            End If
            If (strSplit(4) <> "") Then
                If (strSplit(4) = DateTime.Date) Then
                    If (strSplit(0) = DateTime.Time) Then
                        If Not checkException Then
                            Close #1
                            checkSchedule = True
                        End If
                    End If
                End If
            End If
        End If
    Loop While (strTemp <> "")
    Close #1
End Function
Private Function checkException() As Boolean
    Dim strTemp As String
    Open App.Path & "\except.bak" For Input As #1
    Do
        strTemp = ""
        Line Input #1, strTemp
    Loop While (strTemp <> "")
    Close #1
    MsgBox DateTime.Date
End Function
