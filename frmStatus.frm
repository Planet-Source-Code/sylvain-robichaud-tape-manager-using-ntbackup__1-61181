VERSION 5.00
Begin VB.Form frmStatus 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Status"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6210
   Icon            =   "frmStatus.frx":0000
   LinkTopic       =   "Status"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   6210
   Begin VB.CommandButton Command2 
      Caption         =   "Stop Query"
      Height          =   975
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3480
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF0000&
      Caption         =   "Start Query"
      Height          =   975
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3480
      Width           =   2895
   End
   Begin VB.ListBox lstStatus 
      Height          =   3375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6255
   End
   Begin VB.Timer Timer1 
      Interval        =   990
      Left            =   5280
      Top             =   3720
   End
End
Attribute VB_Name = "frmStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private TapeHandle As Long

Private Sub Command1_Click()
    Timer1.Enabled = True
    Command1.BackColor = vbBlue
    Command2.BackColor = &H8000000F
End Sub

Private Sub Command2_Click()
    Command1.BackColor = &H8000000F
    Command2.BackColor = vbBlue
    Timer1.Enabled = False
    CloseHandle (TapeHandle)
End Sub

Private Sub Form_Load()
    Dim secatt As SECURITY_ATTRIBUTES
    secatt.bInheritHandle = 0&
    secatt.lpSecurityDescriptor = 0&
    secatt.nLength = 0&
    TapeHandle = CreateFile("\\.\Tape0", GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, secatt, OPEN_EXISTING, 0, 0&)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CloseHandle (TapeHandle)
End Sub

Private Sub Timer1_Timer()
    Dim lngRet As Long
    lngRet = GetTapeStatus(ByVal TapeHandle)
    Select Case lngRet
        Case 19:
            lstStatus.AddItem ("The media is write protected.")
        Case 21:
            lstStatus.AddItem ("Loading/Unloading tape")
        Case 50:
            lstStatus.AddItem ("The tape driver does not support a requested function.")
        Case 1102:
            lstStatus.AddItem ("An attempt to access data before the beginning-of-medium marker failed.")
        Case 1111:
            lstStatus.AddItem ("A reset condition was detected on the bus")
        Case 1107:
            lstStatus.AddItem ("The partition information could not be found when a tape was being loaded.")
        Case 1165:
            lstStatus.AddItem ("The tape drive is capable of reporting that it requires cleaning, and reports that it does require cleaning.")
        Case 1100:
            lstStatus.AddItem ("The end-of-tape marker was reached during an operation")
        Case 1101:
            lstStatus.AddItem ("A filemark was reached during an operation.")
        Case 1106:
            lstStatus.AddItem ("The block size is incorrect on a new tape in a multivolume partition.")
        Case 1110:
            lstStatus.AddItem ("The tape that was in the drive has been replaced or removed")
        Case 1104:
            lstStatus.AddItem ("The end-of-data marker was reached during an operation")
        Case 1112:
            lstStatus.AddItem ("There is no media in the drive.")
        Case 1105:
            lstStatus.AddItem ("The tape could not be partitioned.")
        Case 1103:
            lstStatus.AddItem ("A setmark was reached during an operation.")
        Case 1108:
            lstStatus.AddItem ("An attempt to lock the ejection mechanism failed.")
        Case 1109:
            lstStatus.AddItem ("An attempt to unload the tape failed.")
        Case 6:
            lstStatus.AddItem ("No Tape engin found, maybe its unplug, or not install, or another program using it now.")
        Case 0:
            lstStatus.AddItem ("Tape loaded and ready")
            Timer1.Enabled = False
            Command1.BackColor = &H8000000F
            Command2.BackColor = vbBlue
            CloseHandle (TapeHandle)
        Case Else
            lstStatus.AddItem (lngRet & ":Unknow status number")
        End Select
        lstStatus.TopIndex = lstStatus.ListCount - 1
        DoEvents
End Sub
