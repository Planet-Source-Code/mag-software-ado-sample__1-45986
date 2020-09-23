VERSION 5.00
Begin VB.Form frmLegend 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   1035
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5280
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1035
   ScaleWidth      =   5280
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraLegenda 
      Caption         =   "Legend"
      Height          =   945
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5235
      Begin VB.Label lblNotAllowed 
         BackColor       =   &H008080FF&
         Height          =   255
         Left            =   210
         TabIndex        =   6
         Top             =   390
         Width           =   405
      End
      Begin VB.Label lblCustom 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3180
         TabIndex        =   5
         Top             =   390
         Width           =   405
      End
      Begin VB.Label lblRequired 
         BackColor       =   &H0000FF00&
         Height          =   255
         Left            =   1800
         TabIndex        =   4
         Top             =   390
         Width           =   405
      End
      Begin VB.Label lblLegenda 
         AutoSize        =   -1  'True
         Caption         =   "Not allowed"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   720
         TabIndex        =   3
         Top             =   390
         Width           =   975
      End
      Begin VB.Label lblLegenda 
         AutoSize        =   -1  'True
         Caption         =   "Required"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   2310
         TabIndex        =   2
         Top             =   390
         Width           =   720
      End
      Begin VB.Label lblLegenda 
         AutoSize        =   -1  'True
         Caption         =   "Custom entry"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   2
         Left            =   3690
         TabIndex        =   1
         Top             =   390
         Width           =   1110
      End
   End
End
Attribute VB_Name = "frmLegend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Load()
'update legend colors
lblRequired.BackColor = CLR_YELLOW
lblCustom.BackColor = CLR_GREEN
lblNotAllowed.BackColor = CLR_RED
End Sub
