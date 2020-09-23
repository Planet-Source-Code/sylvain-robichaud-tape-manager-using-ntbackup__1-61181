VERSION 5.00
Begin VB.Form frmCalendar 
   Caption         =   "Calendar"
   ClientHeight    =   8400
   ClientLeft      =   19455
   ClientTop       =   825
   ClientWidth     =   8265
   Icon            =   "frmCalendar.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8400
   ScaleWidth      =   8265
   Begin VB.CommandButton cmdWeekHor 
      Height          =   1095
      Index           =   5
      Left            =   120
      TabIndex        =   56
      Top             =   6960
      Width           =   495
   End
   Begin VB.CommandButton cmdDay 
      Caption         =   "Command1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   41
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   63
      Top             =   6960
      Width           =   1095
   End
   Begin VB.CommandButton cmdDay 
      Caption         =   "Command1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   40
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   62
      Top             =   6960
      Width           =   1095
   End
   Begin VB.CommandButton cmdDay 
      Caption         =   "Command1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   39
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   61
      Top             =   6960
      Width           =   1095
   End
   Begin VB.CommandButton cmdDay 
      Caption         =   "Command1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   38
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   60
      Top             =   6960
      Width           =   1095
   End
   Begin VB.CommandButton cmdDay 
      Caption         =   "Command1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   37
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   59
      Top             =   6960
      Width           =   1095
   End
   Begin VB.CommandButton cmdDay 
      Caption         =   "Command1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   36
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   58
      Top             =   6960
      Width           =   1095
   End
   Begin VB.CommandButton cmdDay 
      Caption         =   "Command1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   35
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   57
      Top             =   6960
      Width           =   1095
   End
   Begin VB.CommandButton cmdWeekVer 
      Height          =   495
      Index           =   6
      Left            =   7080
      TabIndex        =   55
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton cmdWeekVer 
      Height          =   495
      Index           =   5
      Left            =   6000
      TabIndex        =   54
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton cmdWeekVer 
      Height          =   495
      Index           =   4
      Left            =   4920
      TabIndex        =   53
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton cmdWeekVer 
      Height          =   495
      Index           =   3
      Left            =   3840
      TabIndex        =   52
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton cmdWeekVer 
      Height          =   495
      Index           =   2
      Left            =   2760
      TabIndex        =   51
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton cmdWeekVer 
      Height          =   495
      Index           =   1
      Left            =   1680
      TabIndex        =   50
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton cmdWeekVer 
      Height          =   495
      Index           =   0
      Left            =   600
      TabIndex        =   49
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton cmdWeekHor 
      Height          =   1095
      Index           =   4
      Left            =   120
      TabIndex        =   48
      Top             =   5880
      Width           =   495
   End
   Begin VB.CommandButton cmdWeekHor 
      Height          =   1095
      Index           =   3
      Left            =   120
      TabIndex        =   47
      Top             =   4800
      Width           =   495
   End
   Begin VB.CommandButton cmdWeekHor 
      Height          =   1095
      Index           =   2
      Left            =   120
      TabIndex        =   46
      Top             =   3720
      Width           =   495
   End
   Begin VB.CommandButton cmdWeekHor 
      Height          =   1095
      Index           =   1
      Left            =   120
      TabIndex        =   45
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton cmdWeekHor 
      Height          =   1095
      Index           =   0
      Left            =   120
      TabIndex        =   44
      Top             =   1560
      Width           =   495
   End
   Begin VB.ComboBox ComboYear 
      Height          =   315
      Left            =   4440
      TabIndex        =   43
      Text            =   "Combo1"
      Top             =   120
      Width           =   3735
   End
   Begin VB.ComboBox ComboMonth 
      Height          =   315
      Left            =   600
      TabIndex        =   42
      Text            =   "Combo1"
      Top             =   120
      Width           =   3735
   End
   Begin VB.CommandButton cmdWeeks 
      Appearance      =   0  'Flat
      BackColor       =   &H80000009&
      Caption         =   "Command1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   6
      Left            =   7080
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton cmdWeeks 
      Appearance      =   0  'Flat
      BackColor       =   &H80000009&
      Caption         =   "Command1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   6000
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton cmdWeeks 
      Appearance      =   0  'Flat
      BackColor       =   &H80000009&
      Caption         =   "Command1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   4920
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton cmdWeeks 
      Appearance      =   0  'Flat
      BackColor       =   &H80000009&
      Caption         =   "Command1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   3840
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton cmdWeeks 
      Appearance      =   0  'Flat
      BackColor       =   &H80000009&
      Caption         =   "Command1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   2760
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton cmdWeeks 
      Appearance      =   0  'Flat
      BackColor       =   &H80000009&
      Caption         =   "Command1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   1680
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton cmdWeeks 
      Appearance      =   0  'Flat
      BackColor       =   &H80000009&
      Caption         =   "Command1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   600
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton cmdDay 
      Caption         =   "Command1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   34
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   5880
      Width           =   1095
   End
   Begin VB.CommandButton cmdDay 
      Caption         =   "Command1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   33
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   5880
      Width           =   1095
   End
   Begin VB.CommandButton cmdDay 
      Caption         =   "Command1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   32
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   5880
      Width           =   1095
   End
   Begin VB.CommandButton cmdDay 
      Caption         =   "Command1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   31
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   5880
      Width           =   1095
   End
   Begin VB.CommandButton cmdDay 
      Caption         =   "Command1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   30
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   5880
      Width           =   1095
   End
   Begin VB.CommandButton cmdDay 
      Caption         =   "Command1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   29
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   5880
      Width           =   1095
   End
   Begin VB.CommandButton cmdDay 
      Caption         =   "Command1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   28
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   5880
      Width           =   1095
   End
   Begin VB.CommandButton cmdDay 
      Caption         =   "Command1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   27
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton cmdDay 
      Caption         =   "Command1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   26
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton cmdDay 
      Caption         =   "Command1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   25
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton cmdDay 
      Caption         =   "Command1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   24
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton cmdDay 
      Caption         =   "Command1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   23
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton cmdDay 
      Caption         =   "Command1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   22
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton cmdDay 
      Caption         =   "Command1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   21
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton cmdDay 
      Caption         =   "Command1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   20
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton cmdDay 
      Caption         =   "Command1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   19
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton cmdDay 
      Caption         =   "Command1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   18
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton cmdDay 
      Caption         =   "Command1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   17
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton cmdDay 
      Caption         =   "Command1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   16
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton cmdDay 
      Caption         =   "Command1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   15
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton cmdDay 
      Caption         =   "Command1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   14
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton cmdDay 
      Caption         =   "Command1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   13
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmdDay 
      Caption         =   "Command1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   12
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmdDay 
      Caption         =   "Command1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   11
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmdDay 
      Caption         =   "Command1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   10
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmdDay 
      Caption         =   "Command1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   9
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmdDay 
      Caption         =   "Command1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   8
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmdDay 
      Caption         =   "Command1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   7
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmdDay 
      Caption         =   "Command1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   6
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton cmdDay 
      Caption         =   "Command1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   5
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton cmdDay 
      Caption         =   "Command1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   4
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton cmdDay 
      Caption         =   "Command1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   3
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton cmdDay 
      Caption         =   "Command1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   2
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton cmdDay 
      Caption         =   "Command1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   1
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton cmdDay 
      Caption         =   "Command1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   0
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1560
      Width           =   1095
   End
End
Attribute VB_Name = "frmCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim intStart As Integer
Private Sub Label2_Click()

End Sub


Private Sub cmdDay_Click(Index As Integer)
'    Form2.txtDate.Text = DateSerial(Val(ComboYear.Text), ComboMonth.ListIndex + 1, cmdDay(Index).Caption)
End Sub

Private Sub ComboMonth_Click()
    clearDays
    fillDays
End Sub



Private Sub ComboYear_Click()
    clearDays
    fillDays
End Sub

Private Sub Form_Load()
    fillMonths
    fillYears
    fillWeeks
    clearDays
    ComboMonth.ListIndex = Format(DateTime.Now, "mm") - 1
    ComboYear.ListIndex = Format(DateTime.Now, "yyyy") - 1990
    fillDays
End Sub

Private Sub clearDays()
    For a = 0 To cmdDay.Count - 1
        cmdDay(a).Caption = " "
        cmdDay(a).Enabled = False
        cmdDay(a).BackColor = &H8000000F
    Next a
End Sub

Private Sub fillDays()
    Dim i As Integer
    Dim temp As Integer
    intStart = Weekday(DateSerial(Val(ComboYear.Text), ComboMonth.ListIndex + 1, 0)) - 1
    For i = intStart To DateSerial(Val(ComboYear.Text), ComboMonth.ListIndex + 2, 1) - DateSerial(Val(ComboYear.Text), ComboMonth.ListIndex + 1, 1) + intStart - 1
        temp = i - intStart + 1
        cmdDay(i).Caption = temp
        cmdDay(i).Enabled = True
    Next i
    If (DateSerial(Val(ComboYear.Text), ComboMonth.ListIndex + 1, Format(DateTime.Now, "dd")) = DateSerial(Format(DateTime.Now, "yyyy"), Format(DateTime.Now, "mm"), Format(DateTime.Now, "dd"))) Then
        cmdDay(Format(DateTime.Now, "dd") + intStart - 1).BackColor = vbBlue
    End If
End Sub
Private Sub fillMonths()
    ComboMonth.Clear
    ComboMonth.AddItem "January"
    ComboMonth.AddItem "February"
    ComboMonth.AddItem "March"
    ComboMonth.AddItem "April"
    ComboMonth.AddItem "May"
    ComboMonth.AddItem "June"
    ComboMonth.AddItem "July"
    ComboMonth.AddItem "August"
    ComboMonth.AddItem "September"
    ComboMonth.AddItem "October"
    ComboMonth.AddItem "November"
    ComboMonth.AddItem "December"
End Sub
Private Sub fillYears()
    ComboYear.Clear
    For i = 1990 To 2100
    ComboYear.AddItem i
    Next i
End Sub
Private Sub fillWeeks()
    cmdWeeks(0).Caption = "Sun"
    cmdWeeks(1).Caption = "Mon"
    cmdWeeks(2).Caption = "Tue"
    cmdWeeks(3).Caption = "Wed"
    cmdWeeks(4).Caption = "Thu"
    cmdWeeks(5).Caption = "Fri"
    cmdWeeks(6).Caption = "Sat"
End Sub

