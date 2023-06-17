VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5475
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12930
   LinkTopic       =   "Form1"
   ScaleHeight     =   5475
   ScaleWidth      =   12930
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Data1 
      Caption         =   "Student Details"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4200
      Visible         =   0   'False
      Width           =   10215
   End
   Begin VB.CommandButton CmdPrevious 
      Caption         =   "&Previous Record"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   3
      Left            =   3000
      TabIndex        =   9
      Top             =   2880
      Width           =   3135
   End
   Begin VB.CommandButton CmdNext 
      Caption         =   "&Next Record"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   2
      Left            =   6360
      TabIndex        =   8
      Top             =   2880
      Width           =   2775
   End
   Begin VB.CommandButton CmdLast 
      Caption         =   "&Last Record"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   9360
      TabIndex        =   7
      Top             =   2880
      Width           =   3255
   End
   Begin VB.CommandButton CmdFirst 
      Caption         =   "&First Record"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   2880
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Index           =   2
      Left            =   4200
      TabIndex        =   5
      Top             =   1800
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Index           =   1
      Left            =   4200
      TabIndex        =   3
      Top             =   960
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Index           =   0
      Left            =   4200
      TabIndex        =   1
      Top             =   120
      Width           =   3495
   End
   Begin VB.Label Label1 
      Caption         =   "Age :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   3615
   End
   Begin VB.Label Label1 
      Caption         =   "Roll Number :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   3615
   End
   Begin VB.Label Label1 
      Caption         =   "Name of Student :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   3615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdFirst_Click(Index As Integer)
DataStudent.Recordest.MoveFirst
End Sub
Private Sub CmdLast_Click(Index As Integer)
DataStudent.Recordest.MoveLast
End Sub

Private Sub CmdNext_Click(Index As Integer)
DataStudent.Recordest.MoveNext
    If DataStudent.Recordest.EOF = True Then
    DataStudent.Recordest.MovePrevious
    MsgBox "This is Last Record.", vbInformation, "Last Record"
    End If
End Sub

Private Sub CmdPrevious_Click(Index As Integer)
DataStudent.Recordest.MovePrevious
    If DataStudent.Recordest.BOF = True Then
    DataStudent.Recordest.MoveNext
    MsgBox "This is Last Record.", vbInformation, "First Record"
    End If
End Sub

