VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmProducts 
   Caption         =   "Products"
   ClientHeight    =   5520
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6825
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5520
   ScaleWidth      =   6825
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtSupplierID 
      DataField       =   "SupplierID"
      DataMember      =   "Products"
      DataSource      =   "deABCStores"
      Height          =   315
      Left            =   1815
      TabIndex        =   22
      Top             =   870
      Width           =   495
   End
   Begin MSComctlLib.StatusBar staProducts 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   21
      Top             =   5025
      Width           =   6825
      _ExtentX        =   12039
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "2/27/02"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "11:42 AM"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdLast 
      Caption         =   "&Last"
      Height          =   495
      Left            =   4560
      TabIndex        =   20
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next"
      Height          =   495
      Left            =   3360
      TabIndex        =   19
      Top             =   3720
      Width           =   975
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "&Previous"
      Height          =   495
      Left            =   1920
      TabIndex        =   18
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmdFirst 
      Caption         =   "&First"
      Height          =   495
      Left            =   480
      TabIndex        =   17
      Top             =   3720
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      DataField       =   "Reorder#"
      DataMember      =   "Products"
      DataSource      =   "deABCStores"
      Height          =   285
      Left            =   1815
      TabIndex        =   16
      Top             =   3195
      Width           =   660
   End
   Begin VB.TextBox txtQtyOnHand 
      DataField       =   "QtyOnHand"
      DataMember      =   "Products"
      DataSource      =   "deABCStores"
      Height          =   285
      Left            =   1815
      TabIndex        =   14
      Top             =   2805
      Width           =   660
   End
   Begin VB.TextBox txtSalePrice 
      DataField       =   "SalePrice"
      DataMember      =   "Products"
      DataSource      =   "deABCStores"
      Height          =   285
      Left            =   1815
      TabIndex        =   12
      Top             =   2430
      Width           =   1320
   End
   Begin VB.TextBox txtCost 
      DataField       =   "Cost"
      DataMember      =   "Products"
      DataSource      =   "deABCStores"
      Height          =   285
      Left            =   1815
      TabIndex        =   10
      Top             =   2055
      Width           =   1320
   End
   Begin VB.TextBox txtDailyRentalPrice 
      DataField       =   "Daily Rental Price"
      DataMember      =   "Products"
      DataSource      =   "deABCStores"
      Height          =   285
      Left            =   1815
      TabIndex        =   8
      Top             =   1665
      Width           =   1320
   End
   Begin VB.CheckBox chkRental 
      DataField       =   "Rental"
      DataMember      =   "Products"
      DataSource      =   "deABCStores"
      Height          =   285
      Left            =   1815
      TabIndex        =   6
      Top             =   1290
      Width           =   330
   End
   Begin VB.TextBox txtDescription 
      DataField       =   "Description"
      DataMember      =   "Products"
      DataSource      =   "deABCStores"
      Height          =   285
      Left            =   1815
      TabIndex        =   3
      Top             =   525
      Width           =   3375
   End
   Begin VB.TextBox txtProductID 
      DataField       =   "ProductID"
      DataMember      =   "Products"
      DataSource      =   "deABCStores"
      Height          =   285
      Left            =   1815
      TabIndex        =   1
      Top             =   150
      Width           =   660
   End
   Begin MSDataListLib.DataCombo dbcSuppliers 
      Bindings        =   "frmProducts.frx":0000
      DataField       =   "SupplierID"
      DataMember      =   "Products"
      DataSource      =   "deABCStores"
      Height          =   315
      Left            =   2520
      TabIndex        =   23
      Top             =   870
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "SupplierName"
      BoundColumn     =   "SupplierID"
      Text            =   ""
      Object.DataMember      =   "Suppliers"
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Reorder &#:"
      Height          =   195
      Index           =   8
      Left            =   1020
      TabIndex        =   15
      Top             =   3240
      Width           =   765
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "&Qty On Hand:"
      Height          =   195
      Index           =   7
      Left            =   810
      TabIndex        =   13
      Top             =   2850
      Width           =   975
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Sale Pri&ce:"
      Height          =   195
      Index           =   6
      Left            =   1020
      TabIndex        =   11
      Top             =   2475
      Width           =   765
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "C&ost:"
      Height          =   195
      Index           =   5
      Left            =   1425
      TabIndex        =   9
      Top             =   2100
      Width           =   360
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Dail&y Rental Price:"
      Height          =   195
      Index           =   4
      Left            =   480
      TabIndex        =   7
      Top             =   1710
      Width           =   1305
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "&Rental:"
      Height          =   195
      Index           =   3
      Left            =   1305
      TabIndex        =   5
      Top             =   1335
      Width           =   510
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "S&upplier ID:"
      Height          =   195
      Index           =   2
      Left            =   960
      TabIndex        =   4
      Top             =   960
      Width           =   825
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Desc&ription:"
      Height          =   195
      Index           =   1
      Left            =   945
      TabIndex        =   2
      Top             =   570
      Width           =   840
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Product &ID:"
      Height          =   195
      Index           =   0
      Left            =   975
      TabIndex        =   0
      Top             =   195
      Width           =   810
   End
End
Attribute VB_Name = "frmProducts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project   : ABC Stores
'Module    : frmProducts
'Date      : 3/2002
'Programmer: Erik Minkin

Option Explicit

Private Sub cmdFirst_Click()
    'Move to the first record
  On Error Resume Next
  deABCStores.rsProducts.MoveFirst
End Sub

Private Sub cmdLast_Click()
    'Move to the last record
  On Error Resume Next
  deABCStores.rsProducts.MoveLast
End Sub

Private Sub cmdNext_Click()
    'Move to the next record
  On Error Resume Next
  With deABCStores.rsProducts
    .MoveNext
    If (.EOF) Then
      .MoveLast
    End If
  End With
End Sub

Private Sub cmdPrevious_Click()
     'Move to the Previous record
  On Error Resume Next
  With deABCStores.rsProducts
    .MovePrevious
    If (.BOF) Then
      .MoveFirst
    End If
  End With
End Sub

