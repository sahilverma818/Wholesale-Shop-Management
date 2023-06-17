VERSION 5.00
Begin VB.Form FrmCustomer 
   Caption         =   "Customer Form"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   Picture         =   "FrmCustomer.frx":0000
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Data DataMobNo 
      Caption         =   "Mobile No"
      Connect         =   "Access"
      DatabaseName    =   "D:\AC Cold Drink Shop\Acshop..s.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   420
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Customer"
      Top             =   1080
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Data DataCustID 
      Caption         =   "Cust_ID"
      Connect         =   "Access"
      DatabaseName    =   "D:\AC Cold Drink Shop\Acshop..s.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   420
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   600
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.CommandButton CmdAllCustomerRecord 
      BackColor       =   &H000080FF&
      Caption         =   "All Customer"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   13320
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4320
      Width           =   2775
   End
   Begin VB.CommandButton CmdSearchRecord 
      BackColor       =   &H000080FF&
      Caption         =   "&Search"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11160
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4320
      Width           =   2055
   End
   Begin VB.CommandButton CmdUpdateRecord 
      BackColor       =   &H000080FF&
      Caption         =   "&Update"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   15600
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3600
      Width           =   2295
   End
   Begin VB.CommandButton CmdCancelRecord 
      BackColor       =   &H000080FF&
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   13320
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3600
      Width           =   2175
   End
   Begin VB.TextBox TxtCustomerID 
      BackColor       =   &H0080FF80&
      DataField       =   "Cust_ID"
      DataSource      =   "DataCustomer"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   13440
      TabIndex        =   0
      Top             =   120
      Width           =   3855
   End
   Begin VB.CommandButton CmdDeleteRecord 
      BackColor       =   &H000080FF&
      Caption         =   "&Delete"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4320
      Width           =   2055
   End
   Begin VB.CommandButton CmdAddNewRecord 
      BackColor       =   &H000080FF&
      Caption         =   "&Add New"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3600
      Width           =   2055
   End
   Begin VB.CommandButton CmdSaveRecord 
      BackColor       =   &H000080FF&
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11160
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3600
      Width           =   2055
   End
   Begin VB.Data DataCustomer 
      Caption         =   "Customer Details"
      Connect         =   "Access"
      DatabaseName    =   "D:\AC Cold Drink Shop\Acshop..s.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Customer"
      Top             =   120
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.CommandButton CmdHome 
      BackColor       =   &H000080FF&
      Caption         =   "Go To &Home"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11880
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   5040
      Width           =   2775
   End
   Begin VB.CommandButton CmdPreviousRecord 
      BackColor       =   &H000080FF&
      Caption         =   "&Previous"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11160
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2880
      Width           =   2055
   End
   Begin VB.TextBox TxtMobno 
      BackColor       =   &H0080FF80&
      DataField       =   "Mob_No"
      DataSource      =   "DataCustomer"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   14160
      TabIndex        =   3
      Top             =   1920
      Width           =   3135
   End
   Begin VB.CommandButton CmdFirstRecord 
      BackColor       =   &H000080FF&
      Caption         =   "&First"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2880
      Width           =   2055
   End
   Begin VB.CommandButton CmdLastRecord 
      BackColor       =   &H000080FF&
      Caption         =   "&Last"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   15600
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2880
      Width           =   2295
   End
   Begin VB.CommandButton CmdExit 
      BackColor       =   &H000080FF&
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   16200
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4320
      Width           =   1695
   End
   Begin VB.CommandButton CmdNextRecord 
      BackColor       =   &H000080FF&
      Caption         =   "&Next"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   13320
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2880
      Width           =   2175
   End
   Begin VB.TextBox TxtCity 
      BackColor       =   &H0080FF80&
      DataField       =   "City"
      DataSource      =   "DataCustomer"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   13440
      TabIndex        =   2
      Top             =   1320
      Width           =   3855
   End
   Begin VB.TextBox TxtCustomerName 
      BackColor       =   &H0080FF80&
      DataField       =   "Cust_Name"
      DataSource      =   "DataCustomer"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   13440
      TabIndex        =   1
      Top             =   720
      Width           =   3855
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080FF80&
      Caption         =   "+91"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13440
      TabIndex        =   21
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FF80&
      Caption         =   "Customer Name"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   9240
      TabIndex        =   20
      Top             =   720
      Width           =   4095
   End
   Begin VB.Label LblMobNo 
      BackColor       =   &H0080FF80&
      Caption         =   "Mob.No. of Customer :"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9240
      TabIndex        =   19
      Top             =   1920
      Width           =   4095
   End
   Begin VB.Label LblCustomerID 
      BackColor       =   &H0080FF80&
      Caption         =   "Customer ID :"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   9240
      TabIndex        =   18
      Top             =   120
      Width           =   4095
   End
   Begin VB.Label LblCity 
      BackColor       =   &H0080FF80&
      Caption         =   "City of Customer :"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9240
      TabIndex        =   16
      Top             =   1320
      Width           =   4095
   End
End
Attribute VB_Name = "FrmCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub MaxFalse()
CmdFirstRecord.Enabled = False
CmdPreviousRecord.Enabled = False
CmdNextRecord.Enabled = False
CmdLastRecord.Enabled = False
CmdAddNewRecord.Enabled = False
CmdSaveRecord.Enabled = True
CmdCancelRecord.Enabled = True
CmdUpdateRecord.Enabled = False
CmdDeleteRecord.Enabled = False
CmdSearchRecord.Enabled = False
CmdAllCustomerRecord.Enabled = False
CmdExit.Enabled = False
CmdHome.Enabled = False
TxtCustomerID.Locked = False
TxtCustomerName.Locked = False
TxtCity.Locked = False
TxtMobno.Locked = False
TxtCustomerID.SetFocus
End Sub
Public Sub MaxTrue()
CmdFirstRecord.Enabled = True
CmdPreviousRecord.Enabled = True
CmdNextRecord.Enabled = True
CmdLastRecord.Enabled = True
CmdAddNewRecord.Enabled = True
CmdSaveRecord.Enabled = False
CmdCancelRecord.Enabled = False
CmdUpdateRecord.Enabled = True
CmdDeleteRecord.Enabled = True
CmdSearchRecord.Enabled = True
CmdAllCustomerRecord.Enabled = TrueH
CmdExit.Enabled = True
CmdHome.Enabled = True
TxtCustomerID.Locked = True
TxtCustomerName.Locked = True
TxtCity.Locked = True
TxtMobno.Locked = True
End Sub
Private Sub CmdFirstRecord_Click()
DataCustomer.Recordset.MoveFirst
End Sub
Private Sub Cmdpreviousrecord_Click()
DataCustomer.Recordset.MovePrevious
    If DataCustomer.Recordset.BOF = True Then
    DataCustomer.Recordset.MoveNext
    MsgBox "This is First Record.", vbInformation, "Frist Record"
    End If
End Sub
Private Sub CmdNextRecord_Click()
DataCustomer.Recordset.MoveNext
    If DataCustomer.Recordset.EOF = True Then
    DataCustomer.Recordset.MovePrevious
    MsgBox "This is Last Record.", vbInformation, "Last Record"
    End If
End Sub
Private Sub CmdLastRecord_Click()
DataCustomer.Recordset.MoveLast
End Sub
Private Sub CmdAddNewRecord_Click()
DataCustomer.Recordset.AddNew
Call MaxFalse
End Sub
Private Sub CmdSaveRecord_Click()
TxtCustomerID.Text = Trim(TxtCustomerID.Text)
TxtCustomerName.Text = Trim(TxtCustomerName.Text)
TxtCity.Text = Trim(TxtCity.Text)
TxtMobno.Text = Trim(TxtMobno.Text)
If TxtCustomerID.Text = "" Then
MsgBox "Please enter a Valid customer ID.", vbCritical, "Blank ID"
TxtCustomerID.SetFocus
ElseIf TxtCustomerName.Text = "" Then
MsgBox "Please enter a Valid Customer Name.", vbCritical, "Blank Nmae"
TxtCustomerName.SetFocus
ElseIf Len(TxtCustomerName.Text) > 30 Then
MsgBox "Please enter a Valid Customer Name maximum 30 characters long only.", vbCritical, "Long Name"
TxtCustomerName.SetFocus
ElseIf TxtCity.Text = "" Or (UCase(TxtCity.Text) <> "SAHARSA" And LCase(TxtCity.Text) <> "supaul" And LCase(TxtCity.Text) <> "madhepura") Then
MsgBox "Please enter City is either -Saharsa,Supaul or Madhepura only.", vbCritical, "Invalid City"
TxtCity.Text = ""
TxtCity.SetFocus
ElseIf TxtMobno.Text = "" Then
MsgBox "Please enter valid Mobile Number.", vbCritical, "Blank Mobile Number"
TxtMob_No.SetFocus
ElseIf Len(TxtMobno.Text) < 10 Or Len(TxtMobno.Text) > 10 Then
MsgBox "Please put mobile number upto 10 digit only.", vbCritical, "Invalid Mobile Number"
TxtMob_No.SetFocus
ElseIf (DataCustomer.Recordset.EOF = False) Then
DataCustomer.Recordset.Update
Call MaxTrue
Else
DataCustID.RecordSource = "Select Cust_ID from customer where Cust_ID='" & TxtCustomerID.Text & "'"
DataCustID.Refresh
DataMobNo.RecordSource = "Select Mob_No from Customer where Mob_No='" & TxtMobno.Text & " '"
DataMobNo.Refresh
    If DataCustID.Recordset.AbsolutePosition = 0 Then
    MsgBox "The Customer ID is already exist,Please enter a another Customer ID", vbCritical, "Duplicate ID"
    TxtCustomerID.Text = ""
    TxtCustomerID.SetFocus
    ElseIf (DataMobNo.Recordset.AbsolutePosition = 0) Then
    MsgBox "The Customer Mobile Number is already exist, Please enter another Mobile Number", vbCritical, "Duplicate Mobile Number"
    TxtMobno.Text = ""
    TxtMobno.SetFocus
    Else
    DataCustomer.Recordset.Update
    Call MaxTrue
    End If
End If
End Sub
Private Sub CmdCancelRecord_Click()
DataCustomer.Recordset.CancelUpdate
Call MaxTrue
End Sub
Private Sub CmdDeleteRecord_Click()
Msg = MsgBox("Are you sure want to Delete the Record of" & TxtCustomerID.Text & ".", vbYesNo + vbExclamation, "Delete Record")
    If Msg = vbYes Then
    DataCustomer.Recordset.Delete
    DataCustomer.Recordset.MoveNext
    If DataCustomer.Recordset.EOF = True Then
    DataCustomer.Recordset.MovePrevious
    End If
End If
End Sub
Private Sub CmdSearchRecord_Click()
ID = InputBox("Please enter a Customer ID,Which Record you want to Search:->", "ID", "'")
DataCustomer.Recordset.FindFirst ("Cust_ID ='" & ID & "'")
End Sub
Private Sub CmdHome_Click()
FrmHome.Show
Unload Me
End Sub
Private Sub CmdExit_Click()
ExitProject
End Sub

Private Sub CmdUpdateRecord_Click()
DataCustomer.Recordset.Edit
Call MaxFalse
End Sub
