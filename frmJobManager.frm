VERSION 5.00
Begin VB.Form frmJobManager 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Job Manager"
   ClientHeight    =   7275
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9375
   Icon            =   "frmJobManager.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7275
   ScaleWidth      =   9375
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   6720
      Width           =   9015
   End
   Begin VB.CommandButton cmdClearAll 
      Caption         =   "Clear All"
      Height          =   495
      Left            =   6360
      TabIndex        =   8
      Top             =   4320
      Width           =   2775
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove"
      Height          =   495
      Left            =   3240
      TabIndex        =   7
      Top             =   4320
      Width           =   2895
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   4320
      Width           =   2775
   End
   Begin VB.FileListBox File1 
      Height          =   2235
      Left            =   5280
      MultiSelect     =   2  'Extended
      TabIndex        =   5
      Top             =   2040
      Width           =   3855
   End
   Begin VB.DirListBox Dir1 
      Height          =   1890
      Left            =   120
      TabIndex        =   4
      Top             =   2400
      Width           =   5175
   End
   Begin VB.ListBox lstPath 
      Height          =   1815
      Left            =   120
      MultiSelect     =   2  'Extended
      TabIndex        =   3
      Top             =   4800
      Width           =   9015
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   5175
   End
   Begin VB.TextBox txtJobName 
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Job Name:"
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   480
      Width           =   855
   End
End
Attribute VB_Name = "frmJobManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim blChange As Boolean
Private Sub cmdAdd_Click()
Dim blpass As Boolean
    For a = 0 To File1.ListCount - 1
        If (File1.Selected(a)) Then
            intelliWrite File1.Path & "\" & File1.List(a)
            blpass = True
        End If
    Next a
    If (blpass = False) Then
        intelliWrite Dir1.List(Dir1.ListIndex)
    End If
End Sub

Private Sub cmdClearAll_Click()
    Dim lngcount As Long
    lngcount = lstPath.ListCount
    lstPath.Clear
    If (lngcount <> lstPath.ListCount) Then
        blChange = True
    End If
End Sub

Private Sub cmdRemove_Click()
    Dim lngcount As Long
    lngcount = lstPath.ListCount
    lstPath.RemoveItem (lstPath.ListIndex)
    If (lngcount <> lstPath.ListCount) Then
        blChange = True
    End If
End Sub

Private Sub Command1_Click()
    Open App.Path & "\files.bak" For Output As #1
    For a = 0 To lstPath.ListCount - 1
        Print #1, lstPath.List(a)
    Next a
    Close #1
    Open App.Path & "\jobs.bak" For Output As #1
    Print #1, txtJobName.Text
    Close #1
    blChange = False
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub

Private Sub intelliWrite(strTemp As String)
    Dim blAleady As Boolean
    Dim lngcount As Long
    lngcount = lstPath.ListCount
    For a = 0 To lstPath.ListCount - 1
        If strTemp = lstPath.List(a) Then
            blAlready = True
        End If
    Next a
    If Not (blAlready) Then
        lstPath.AddItem (strTemp)
    End If
    If lngcount <> lstPath.ListCount Then
        blChange = True
    End If
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Dim strTemp As String
    Me.Left = 0
    Me.Top = 0
    blChange = False
    Open App.Path & "\files.bak" For Input As #1
    lstPath.Clear
    Do
        strTemp = ""
        Line Input #1, strTemp
        If (strTemp <> "") Then
            lstPath.AddItem (strTemp)
        End If
    Loop While (strTemp <> "")
    Close #1
    Open App.Path & "\jobs.bak" For Input As #1
    Line Input #1, strTemp
    If (strTemp <> "") Then
        txtJobName.Text = strTemp
    End If
    Close #1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If blChange Then
        If (MsgBox("Do you want to save the file's list ?", vbYesNo, "Save") = vbYes) Then
            Open App.Path & "\files.bak" For Output As #1
            For a = 0 To lstPath.ListCount - 1
                Print #1, lstPath.List(a)
            Next a
            Close #1
            Open App.Path & "\jobs.bak" For Output As #1
            Print #1, txtJobName.Text
            Close #1
        End If
    End If
End Sub
