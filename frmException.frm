VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form frmException 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exception (Days without running backup)"
   ClientHeight    =   6150
   ClientLeft      =   420
   ClientTop       =   1380
   ClientWidth     =   8850
   Icon            =   "frmException.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   8850
   Begin VB.CommandButton cmdDone 
      Caption         =   "Save"
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   5520
      Width           =   8415
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear All"
      Height          =   375
      Left            =   4560
      TabIndex        =   3
      Top             =   5040
      Width           =   4095
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "Remove"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   5040
      Width           =   4215
   End
   Begin VB.ListBox lstExcept 
      Height          =   1815
      Left            =   240
      TabIndex        =   1
      Top             =   3120
      Width           =   8415
   End
   Begin MSACAL.Calendar Calendar1 
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   8415
      _Version        =   524288
      _ExtentX        =   14843
      _ExtentY        =   4895
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
End
Attribute VB_Name = "frmException"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim blSave As Boolean
Dim blChange As Boolean
Private Sub Calendar1_Click()
    Dim blCheck As Boolean
    Dim dtDate As Date
    Dim lngcount As Long
    lngcount = lstExcept.ListCount
    For a = 0 To lstExcept.ListCount - 1
        If (Calendar1.Value = lstExcept.List(a)) Then
            blCheck = True
        End If
    Next a
    If blCheck <> True Then
        lstExcept.AddItem Calendar1.Value
        For a = 0 To lstExcept.ListCount - 1
            For b = a To lstExcept.ListCount - 1
                If CDate(lstExcept.List(a)) > CDate(lstExcept.List(b)) Then
                    dtDate = lstExcept.List(a)
                    lstExcept.List(a) = lstExcept.List(b)
                    lstExcept.List(b) = dtDate
                End If
            Next b
        Next a
    End If
    If (lngcount <> lstExcept.ListCount) Then
        blChange = True
    End If
End Sub

Private Sub cmdClear_Click()
    Dim lngcount As Long
    lngcount = lstExcept.ListCount
    lstExcept.Clear
    If lngcount <> lstExcept.ListCount Then
        blChange = True
    End If
End Sub

Private Sub cmdDel_Click()
    Dim lngcount As Long
    lngcount = lstExcept.ListCount
    lstExcept.RemoveItem lstExcept.ListIndex
    If lngcount <> lstExcept.ListCount Then
        blChange = True
    End If
End Sub



Private Sub cmdDone_Click()
    blSave = True
    Open App.Path & "\except.bak" For Output As #1
    For a = 0 To lstExcept.ListCount - 1
        Print #1, lstExcept.List(a)
    Next a
    Close #1
    blChange = False
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Dim strTemp As String
    Me.Left = 0
    Me.Top = 0
    blChange = False
    Open App.Path & "\except.bak" For Input As #1
    lstExcept.Clear
    Do
        strTemp = ""
        Line Input #1, strTemp
        If (strTemp <> "") Then
            lstExcept.AddItem (strTemp)
        End If
    Loop While (strTemp <> "")
    Close #1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not blSave And blChange Then
        If (MsgBox("Do you want to save the exception's list ?", vbYesNo, "Save") = vbYes) Then
            Open App.Path & "\except.bak" For Output As #1
            For a = 0 To lstExcept.ListCount - 1
                Print #1, lstExcept.List(a)
            Next a
            Close #1
        End If
    End If
End Sub
