VERSION 5.00
Begin VB.Form FrmCust_Details 
   Caption         =   "Customer Details"
   ClientHeight    =   7470
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14475
   LinkTopic       =   "Form1"
   ScaleHeight     =   7470
   ScaleWidth      =   14475
   StartUpPosition =   3  'Windows Default
   WindowState     =   1  'Minimized
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "D:\AC Cold Drink Shop\Acshop..s.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   1080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Sales"
      Top             =   4560
      Width           =   3135
   End
End
Attribute VB_Name = "FrmCust_Details"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
