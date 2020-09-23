VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form frmScheduler 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Scheduler"
   ClientHeight    =   6825
   ClientLeft      =   420
   ClientTop       =   1200
   ClientWidth     =   8625
   Icon            =   "frmScheduler.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   8625
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   495
      Left            =   120
      TabIndex        =   41
      Top             =   6240
      Width           =   8295
   End
   Begin VB.Frame Frame1 
      Caption         =   "Clock"
      Height          =   855
      Left            =   5280
      TabIndex        =   32
      Top             =   0
      Width           =   3135
      Begin VB.ListBox lstHour 
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   480
         Width           =   615
      End
      Begin VB.ListBox lstMin 
         Height          =   255
         Left            =   840
         TabIndex        =   35
         Top             =   480
         Width           =   615
      End
      Begin VB.ListBox lstSec 
         Height          =   255
         Left            =   1560
         TabIndex        =   34
         Top             =   480
         Width           =   615
      End
      Begin VB.ListBox lstAMPM 
         Height          =   255
         ItemData        =   "frmScheduler.frx":0ECA
         Left            =   2280
         List            =   "frmScheduler.frx":0ED4
         TabIndex        =   33
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label12 
         Caption         =   " :"
         Height          =   255
         Left            =   1440
         TabIndex        =   40
         Top             =   480
         Width           =   135
      End
      Begin VB.Label Label11 
         Caption         =   " :"
         Height          =   255
         Left            =   720
         TabIndex        =   39
         Top             =   480
         Width           =   135
      End
      Begin VB.Label Label8 
         Caption         =   "HH"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label9 
         Caption         =   "MM"
         Height          =   255
         Left            =   840
         TabIndex        =   37
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label10 
         Caption         =   "SS"
         Height          =   255
         Left            =   1560
         TabIndex        =   36
         Top             =   240
         Width           =   255
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Clear All"
      Height          =   375
      Left            =   5880
      TabIndex        =   31
      Top             =   5760
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Add"
      Height          =   375
      Left            =   120
      TabIndex        =   30
      Top             =   5760
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Remove"
      Height          =   375
      Left            =   3000
      TabIndex        =   29
      Top             =   5760
      Width           =   2655
   End
   Begin VB.ListBox lstSchedule 
      Height          =   1425
      Left            =   120
      TabIndex        =   28
      Top             =   4320
      Width           =   8295
   End
   Begin VB.ComboBox cmdSchedule 
      Height          =   315
      ItemData        =   "frmScheduler.frx":0EE0
      Left            =   960
      List            =   "frmScheduler.frx":0EF0
      TabIndex        =   1
      Top             =   240
      Width           =   1815
   End
   Begin VB.Frame fmMonthly 
      Caption         =   "Monthly"
      Height          =   3375
      Left            =   120
      TabIndex        =   16
      Top             =   840
      Visible         =   0   'False
      Width           =   8295
      Begin VB.ListBox lstDays 
         Height          =   1425
         ItemData        =   "frmScheduler.frx":0F12
         Left            =   1920
         List            =   "frmScheduler.frx":0F2B
         TabIndex        =   22
         Top             =   960
         Width           =   1215
      End
      Begin VB.ListBox lstCount 
         Height          =   1035
         ItemData        =   "frmScheduler.frx":0F6F
         Left            =   960
         List            =   "frmScheduler.frx":0F82
         TabIndex        =   21
         Top             =   960
         Width           =   855
      End
      Begin VB.OptionButton Option2 
         Caption         =   "The"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   960
         Width           =   735
      End
      Begin VB.ComboBox cmbMonth 
         Height          =   315
         Left            =   960
         TabIndex        =   18
         Top             =   480
         Width           =   615
      End
      Begin VB.OptionButton Option1 
         Caption         =   "The"
         Height          =   375
         Left            =   240
         TabIndex        =   17
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "of the month"
         Height          =   375
         Left            =   3240
         TabIndex        =   23
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "day of the month"
         Height          =   255
         Left            =   1680
         TabIndex        =   19
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.Frame fmDaily 
      Caption         =   "Daily"
      Height          =   3375
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Visible         =   0   'False
      Width           =   8295
      Begin VB.ComboBox cmbDailyDay 
         Height          =   315
         Left            =   1080
         TabIndex        =   4
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Day(s)"
         Height          =   255
         Left            =   1920
         TabIndex        =   5
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Every:"
         Height          =   255
         Left            =   480
         TabIndex        =   3
         Top             =   720
         Width           =   495
      End
   End
   Begin VB.Frame fmWeekly 
      Caption         =   "Weekly"
      Height          =   3375
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Visible         =   0   'False
      Width           =   8295
      Begin VB.CheckBox chkFriday 
         Caption         =   "Friday"
         Height          =   255
         Left            =   1920
         TabIndex        =   15
         Top             =   2280
         Width           =   1095
      End
      Begin VB.CheckBox chkThursday 
         Caption         =   "Thursday"
         Height          =   375
         Left            =   1920
         TabIndex        =   14
         Top             =   1800
         Width           =   1215
      End
      Begin VB.CheckBox chkWednesday 
         Caption         =   "Wednesday"
         Height          =   255
         Left            =   1920
         TabIndex        =   13
         Top             =   1440
         Width           =   1335
      End
      Begin VB.CheckBox chkTuesday 
         Caption         =   "Tuesday"
         Height          =   375
         Left            =   1920
         TabIndex        =   12
         Top             =   960
         Width           =   975
      End
      Begin VB.CheckBox chkMonday 
         Caption         =   "Monday"
         Height          =   375
         Left            =   1920
         TabIndex        =   11
         Top             =   480
         Width           =   975
      End
      Begin VB.CheckBox chkSunday 
         Caption         =   "Sunday"
         Height          =   375
         Left            =   3720
         TabIndex        =   10
         Top             =   960
         Width           =   1335
      End
      Begin VB.CheckBox chkSaturday 
         Caption         =   "Saturday"
         Height          =   255
         Left            =   3720
         TabIndex        =   9
         Top             =   600
         Width           =   1455
      End
      Begin VB.ComboBox cmbWeekly 
         Height          =   315
         Left            =   840
         TabIndex        =   7
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Every"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   480
         Width           =   495
      End
   End
   Begin VB.Frame fmOnce 
      Caption         =   "Once"
      Height          =   3375
      Left            =   120
      TabIndex        =   24
      Top             =   840
      Visible         =   0   'False
      Width           =   8295
      Begin VB.TextBox txtDate 
         Height          =   375
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   480
         Width           =   2775
      End
      Begin MSACAL.Calendar crOnce 
         Height          =   2175
         Left            =   1200
         TabIndex        =   26
         Top             =   960
         Visible         =   0   'False
         Width           =   4455
         _Version        =   524288
         _ExtentX        =   7858
         _ExtentY        =   3836
         _StockProps     =   1
         BackColor       =   -2147483633
         Year            =   2005
         Month           =   6
         Day             =   10
         DayLength       =   1
         MonthLength     =   2
         DayFontColor    =   0
         FirstDay        =   1
         GridCellEffect  =   1
         GridFontColor   =   10485760
         GridLinesColor  =   -2147483632
         ShowDateSelectors=   -1  'True
         ShowDays        =   -1  'True
         ShowHorizontalGrid=   -1  'True
         ShowTitle       =   -1  'True
         ShowVerticalGrid=   -1  'True
         TitleFontColor  =   10485760
         ValueIsNull     =   0   'False
         BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label7 
         Caption         =   "Run on : "
         Height          =   255
         Left            =   480
         TabIndex        =   25
         Top             =   480
         Width           =   735
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Schedule:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "frmScheduler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim blChange As Boolean

Private Sub cmdSave_Click()
    Open App.Path & "\schedule.bak" For Output As #1
    For a = 0 To lstSchedule.ListCount - 1
        Print #1, lstSchedule.List(a)
    Next a
    Close #1
    blChange = False
End Sub

Private Sub cmdSchedule_Click()
    fmDaily.Visible = False
    fmWeekly.Visible = False
    fmMonthly.Visible = False
    fmOnce.Visible = False
    Select Case cmdSchedule.Text
        Case "Daily"
            fmDaily.Visible = True
        Case "Weekly"
            fmWeekly.Visible = True
        Case "Monthly"
            fmMonthly.Visible = True
        Case "Once"
            fmOnce.Visible = True
    End Select
End Sub

Private Sub Command1_Click()
    Dim lngCount As Long
    lngCount = lstSchedule.ListCount
    lstSchedule.RemoveItem (lstSchedule.ListIndex)
    If (lngCount <> lstSchedule.ListCount) Then
        blChange = True
    End If
End Sub

Private Sub Command2_Click()
    Dim strTime As String
    Dim lngCount As Long
    Dim lngBuffer As Long
    lngCount = lstSchedule.ListCount
    strTime = lstHour.List(lstHour.ListIndex) & ":" & lstMin.List(lstMin.ListIndex) & ":" & lstSec.List(lstSec.ListIndex) & " " & lstAMPM.List(lstAMPM.ListIndex)
    If fmDaily.Visible = True Then
        intelliWrite (strTime & "|" & cmbDailyDay.List(cmbDailyDay.ListIndex) & "@" & DateTime.Date & "|||")
    End If
    If fmWeekly.Visible = True Then
        If (chkSunday.Value = vbChecked) Then
            lngBuffer = DateTime.Weekday(DateTime.Now) - 1
            If (lngBuffer = 0) Then
                intelliWrite (strTime & "||" & cmbWeekly.List(cmbWeekly.ListIndex) & ":" & DateTime.Date & "||")
            Else
                If (lngBuffer < 0) Then
                    lngBuffer = lngBuffer + 7
                End If
                intelliWrite (strTime & "||" & cmbWeekly.List(cmbWeekly.ListIndex) & ":" & DateTime.DateAdd("d", -(lngBuffer), DateTime.Date) & "||")
            End If
        End If
        If (chkMonday.Value = vbChecked) Then
            lngBuffer = DateTime.Weekday(DateTime.Now) - 2
            If (lngBuffer = 0) Then
                intelliWrite (strTime & "||" & cmbWeekly.List(cmbWeekly.ListIndex) & ":" & DateTime.Date & "||")
            Else
                If (lngBuffer < 0) Then
                    lngBuffer = lngBuffer + 7
                End If
                intelliWrite (strTime & "||" & cmbWeekly.List(cmbWeekly.ListIndex) & ":" & DateTime.DateAdd("d", -(lngBuffer), DateTime.Date) & "||")
            End If
        End If
        If (chkTuesday.Value = vbChecked) Then
            lngBuffer = DateTime.Weekday(DateTime.Now) - 3
            If (lngBuffer = 0) Then
                intelliWrite (strTime & "||" & cmbWeekly.List(cmbWeekly.ListIndex) & ":" & DateTime.Date & "||")
            Else
                If (lngBuffer < 0) Then
                    lngBuffer = lngBuffer + 7
                End If
                intelliWrite (strTime & "||" & cmbWeekly.List(cmbWeekly.ListIndex) & ":" & DateTime.DateAdd("d", -(lngBuffer), DateTime.Date) & "||")
            End If
        End If
        If (chkWednesday.Value = vbChecked) Then
            lngBuffer = DateTime.Weekday(DateTime.Now) - 4
            If (lngBuffer = 0) Then
                intelliWrite (strTime & "||" & cmbWeekly.List(cmbWeekly.ListIndex) & ":" & DateTime.Date & "||")
            Else
                If (lngBuffer < 0) Then
                    lngBuffer = lngBuffer + 7
                End If
                intelliWrite (strTime & "||" & cmbWeekly.List(cmbWeekly.ListIndex) & ":" & DateTime.DateAdd("d", -(lngBuffer), DateTime.Date) & "||")
            End If
        End If
        If (chkThursday.Value = vbChecked) Then
            lngBuffer = DateTime.Weekday(DateTime.Now) - 5
            If (lngBuffer = 0) Then
                intelliWrite (strTime & "||" & cmbWeekly.List(cmbWeekly.ListIndex) & ":" & DateTime.Date & "||")
            Else
                If (lngBuffer < 0) Then
                    lngBuffer = lngBuffer + 7
                End If
                intelliWrite (strTime & "||" & cmbWeekly.List(cmbWeekly.ListIndex) & ":" & DateTime.DateAdd("d", -(lngBuffer), DateTime.Date) & "||")
            End If
        End If
        If (chkFriday.Value = vbChecked) Then
            lngBuffer = DateTime.Weekday(DateTime.Now) - 6
            If (lngBuffer = 0) Then
                intelliWrite (strTime & "||" & cmbWeekly.List(cmbWeekly.ListIndex) & ":" & DateTime.Date & "||")
            Else
                If (lngBuffer < 0) Then
                    lngBuffer = lngBuffer + 7
                End If
                intelliWrite (strTime & "||" & cmbWeekly.List(cmbWeekly.ListIndex) & ":" & DateTime.DateAdd("d", -(lngBuffer), DateTime.Date) & "||")
            End If
        End If
        If (chkSaturday.Value = vbChecked) Then
            lngBuffer = DateTime.Weekday(DateTime.Now) - 7
            If (lngBuffer = 0) Then
                intelliWrite (strTime & "||" & cmbWeekly.List(cmbWeekly.ListIndex) & ":" & DateTime.Date & "||")
            Else
                If (lngBuffer < 0) Then
                    lngBuffer = lngBuffer + 7
                End If
                intelliWrite (strTime & "||" & cmbWeekly.List(cmbWeekly.ListIndex) & ":" & DateTime.DateAdd("d", -(lngBuffer), DateTime.Date) & "||")
            End If
        End If
    End If
    If fmOnce.Visible = True Then
        intelliWrite (strTime & "||||" & txtDate.Text)
    End If
    If fmMonthly.Visible = True Then
        If (Option1.Value = True) Then
            intelliWrite (strTime & "|||d:" & cmbMonth.List(cmbMonth.ListIndex) & "|")
        Else
            Dim intVal As Integer
            Dim strString As String
            Select Case lstDays.List(lstDays.ListIndex)
                Case "Sunday"
                       strString = strTime & "|||w:" & lstCount.ListIndex + 1 & ":1"
                Case "Monday"
                    strString = strTime & "|||w:" & lstCount.ListIndex + 1 & ":2"
                Case "Tuesday"
                    strString = strTime & "|||w:" & lstCount.ListIndex + 1 & ":3"
                Case "Wednesday"
                    strString = strTime & "|||w:" & lstCount.ListIndex + 1 & ":4"
                Case "Thursday"
                    strString = strTime & "|||w:" & lstCount.ListIndex + 1 & ":5"
                Case "Friday"
                    strString = strTime & "|||w:" & lstCount.ListIndex + 1 & ":6"
                Case "Saturday"
                    strString = strTime & "|||w:" & lstCount.ListIndex + 1 & ":7"
            End Select
            intelliWrite (strString & "|")
        End If
    End If
    If (lngCount <> lstSchedule.ListCount) Then
        blChange = True
    End If
End Sub

Private Sub Command3_Click()
    Dim lngCount As Long
    lngCount = lstSchedule.ListCount
    lstSchedule.Clear
    If lngCount <> lstSchedule.ListCount Then
        blChange = True
    End If
End Sub


Private Sub crOnce_Click()
    txtDate.Text = crOnce.Value
    crOnce.Visible = False
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Dim strTemp As String
    Me.Left = 0
    Me.Top = 0
    blChange = False
    Open App.Path & "\schedule.bak" For Input As #1
    lstSchedule.Clear
    Do
        strTemp = ""
        Line Input #1, strTemp
        If (strTemp <> "") Then
            lstSchedule.AddItem (strTemp)
        End If
    Loop While (strTemp <> "")
    Close #1
    cmbDailyDay.Clear
    cmbWeekly.Clear
    For a = 1 To 366
        cmbDailyDay.AddItem (a)
    Next a
    For a = 1 To 53
        cmbWeekly.AddItem (a)
    Next a
    For a = 1 To 31
        cmbMonth.AddItem (a)
    Next a
    For a = 12 To 1 Step -1
        lstHour.AddItem (a)
        If (a = DateTime.Hour(DateTime.Now)) Then
            lstHour.Selected(12 - a) = True
        End If
        If ((a + 12) = DateTime.Hour(DateTime.Now)) Then
            lstHour.Selected(12 - a) = True
            lstAMPM.Selected(0) = True
        Else
            lstAMPM.Selected(1) = True
        End If
    Next a
    For a = 59 To 0 Step -1
        If (a < 10) Then
            lstMin.AddItem ("0" & a)
            lstSec.AddItem ("0" & a)
        Else
            lstMin.AddItem (a)
            lstSec.AddItem (a)
        End If
        If (a = DateTime.Minute(DateTime.Now)) Then
            lstMin.Selected(59 - a) = True
        End If
        If (a = DateTime.Second(DateTime.Now)) Then
            lstSec.Selected(59 - a) = True
        End If
    Next a
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If blChange Then
        If (MsgBox("Do you want to save the schedule's list?", vbYesNo, "Save") = vbYes) Then
            Open App.Path & "\schedule.bak" For Output As #1
            For a = 0 To lstSchedule.ListCount - 1
                Print #1, lstSchedule.List(a)
            Next a
            Close #1
        End If
    End If
End Sub
Private Sub lstAMPM_Scroll()
    lstAMPM.Selected(lstAMPM.TopIndex) = True
End Sub

Private Sub lstHour_Scroll()
    lstHour.Selected(lstHour.TopIndex) = True
End Sub

Private Sub lstMin_Scroll()
    lstMin.Selected(lstMin.TopIndex) = True
End Sub
Private Sub lstSec_Scroll()
    lstSec.Selected(lstSec.TopIndex) = True
End Sub

Private Sub txtDate_Click()
    crOnce.Visible = True
End Sub


Private Sub intelliWrite(strTemp As String)
    Dim blAleady As Boolean
    Dim lngCount As Long
    lngCount = lstSchedule.ListCount
    For a = 0 To lstSchedule.ListCount - 1
        If strTemp = lstSchedule.List(a) Then
            blAlready = True
        End If
    Next a
    If Not (blAlready) Then
        lstSchedule.AddItem (strTemp)
    End If
    If lngCount <> lstSchedule.ListCount Then
        blChange = True
    End If
End Sub
