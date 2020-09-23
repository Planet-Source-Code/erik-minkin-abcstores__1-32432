VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmCustomers 
   Caption         =   "Customers"
   ClientHeight    =   6135
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6420
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6135
   ScaleWidth      =   6420
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdADD 
      Caption         =   "&Add"
      CausesValidation=   0   'False
      Height          =   495
      Left            =   5160
      TabIndex        =   27
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Sa&ve"
      Height          =   495
      Left            =   5160
      TabIndex        =   26
      Top             =   3600
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      CausesValidation=   0   'False
      Height          =   495
      Left            =   5160
      TabIndex        =   25
      Top             =   3000
      Width           =   975
   End
   Begin MSComctlLib.StatusBar staCustomers 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   24
      Top             =   5760
      Width           =   6420
      _ExtentX        =   11324
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "3/5/02"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "4:58 PM"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdLast 
      Caption         =   "&Last"
      Height          =   495
      Left            =   4080
      TabIndex        =   23
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next"
      Height          =   495
      Left            =   2760
      TabIndex        =   22
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "&Previous"
      Height          =   495
      Left            =   1560
      TabIndex        =   21
      Top             =   4440
      Width           =   975
   End
   Begin VB.CommandButton cmdFirst 
      Caption         =   "&First"
      Height          =   495
      Left            =   360
      TabIndex        =   20
      Top             =   4440
      Width           =   975
   End
   Begin VB.TextBox txtExpDate 
      DataField       =   "ExpDate"
      DataMember      =   "Customers"
      DataSource      =   "deABCStores"
      Height          =   285
      Left            =   2010
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   3810
      Width           =   1320
   End
   Begin VB.TextBox txtCreditCardNumber 
      DataField       =   "CreditCardNumber"
      DataMember      =   "Customers"
      DataSource      =   "deABCStores"
      Height          =   285
      Left            =   2010
      Locked          =   -1  'True
      MaxLength       =   16
      TabIndex        =   17
      Top             =   3435
      Width           =   2640
   End
   Begin VB.TextBox txtPaymentType 
      DataField       =   "PaymentType"
      DataMember      =   "Customers"
      DataSource      =   "deABCStores"
      Height          =   285
      Left            =   2010
      Locked          =   -1  'True
      MaxLength       =   15
      TabIndex        =   15
      Top             =   3045
      Width           =   2475
   End
   Begin VB.TextBox txtZipCode 
      DataField       =   "ZipCode"
      DataMember      =   "Customers"
      DataSource      =   "deABCStores"
      Height          =   285
      Left            =   2010
      Locked          =   -1  'True
      MaxLength       =   5
      TabIndex        =   13
      Top             =   2670
      Width           =   1815
   End
   Begin VB.TextBox txtState 
      DataField       =   "State"
      DataMember      =   "Customers"
      DataSource      =   "deABCStores"
      Height          =   285
      Left            =   2010
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   11
      Top             =   2295
      Width           =   450
   End
   Begin VB.TextBox txtCity 
      DataField       =   "City"
      DataMember      =   "Customers"
      DataSource      =   "deABCStores"
      Height          =   285
      Left            =   2010
      Locked          =   -1  'True
      MaxLength       =   25
      TabIndex        =   9
      Top             =   1905
      Width           =   2475
   End
   Begin VB.TextBox txtAddress 
      DataField       =   "Address"
      DataMember      =   "Customers"
      DataSource      =   "deABCStores"
      Height          =   285
      Left            =   2010
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1530
      Width           =   3375
   End
   Begin VB.TextBox txtFirstName 
      DataField       =   "FirstName"
      DataMember      =   "Customers"
      DataSource      =   "deABCStores"
      Height          =   285
      Left            =   2010
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   1155
      Width           =   1650
   End
   Begin VB.TextBox txtLastName 
      DataField       =   "LastName"
      DataMember      =   "Customers"
      DataSource      =   "deABCStores"
      Height          =   285
      Left            =   2010
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   765
      Width           =   3300
   End
   Begin VB.TextBox txtCustomerID 
      DataField       =   "CustomerID"
      DataMember      =   "Customers"
      DataSource      =   "deABCStores"
      Height          =   285
      Left            =   2010
      Locked          =   -1  'True
      MaxLength       =   4
      TabIndex        =   1
      Top             =   390
      Width           =   825
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "E&xp Date:"
      Height          =   195
      Index           =   9
      Left            =   1275
      TabIndex        =   18
      Top             =   3855
      Width           =   705
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Credit Card Num&ber:"
      Height          =   195
      Index           =   8
      Left            =   555
      TabIndex        =   16
      Top             =   3480
      Width           =   1425
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Pa&yment Type:"
      Height          =   195
      Index           =   7
      Left            =   915
      TabIndex        =   14
      Top             =   3090
      Width           =   1065
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "&Zip Code:"
      Height          =   195
      Index           =   6
      Left            =   1290
      TabIndex        =   12
      Top             =   2715
      Width           =   690
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Sta&te:"
      Height          =   195
      Index           =   5
      Left            =   1560
      TabIndex        =   10
      Top             =   2340
      Width           =   420
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "&City:"
      Height          =   195
      Index           =   4
      Left            =   1680
      TabIndex        =   8
      Top             =   1950
      Width           =   300
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Addre&ss:"
      Height          =   195
      Index           =   3
      Left            =   1365
      TabIndex        =   6
      Top             =   1575
      Width           =   615
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Fi&rst Name:"
      Height          =   195
      Index           =   2
      Left            =   1185
      TabIndex        =   4
      Top             =   1200
      Width           =   795
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Last Nam&e:"
      Height          =   195
      Index           =   1
      Left            =   1170
      TabIndex        =   2
      Top             =   810
      Width           =   810
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "C&ustomer ID:"
      Height          =   195
      Index           =   0
      Left            =   1065
      TabIndex        =   0
      Top             =   435
      Width           =   915
   End
End
Attribute VB_Name = "frmCustomers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project   : ABC Stores
'Module    : frmCustomer
'Date      : 3/2002
'Programmer: Erik Minkin

Option Explicit

Private Sub setUpAdd()
    'Set up controls for the add
    With txtCustomerID
      .Locked = False
      .SetFocus
    End With
    cmdADD.Caption = "&Cancel"
    cmdSave.Enabled = True
    frmMDIParent.tlbABC.Enabled = False
End Sub

Private Sub EnableControls()
   'Enable navigation buttons
  cmdNext.Enabled = True
  cmdPrevious.Enabled = True
  cmdFirst.Enabled = True
  cmdLast.Enabled = True
  cmdDelete.Enabled = True
  frmMDIParent.tlbABC.Enabled = True
  frmMDIParent.mnuForm.Enabled = True
  frmMDIParent.mnuHelp.Enabled = True
  frmMDIParent.mnuReport.Enabled = True
  frmMDIParent.mnuWindow.Enabled = True
End Sub

Private Sub DisableControls()
   'Disable navigation buttons
  cmdNext.Enabled = False
  cmdPrevious.Enabled = False
  cmdFirst.Enabled = False
  cmdLast.Enabled = False
  cmdDelete.Enabled = False
  frmMDIParent.mnuForm.Enabled = False
  frmMDIParent.mnuHelp.Enabled = False
  frmMDIParent.mnuReport.Enabled = False
  frmMDIParent.mnuWindow.Enabled = False
End Sub

Private Sub DisableTextBoxes()
  txtCustomerID.Locked = True
  txtAddress.Locked = True
  txtCity.Locked = True
  txtCreditCardNumber.Locked = True
  txtExpDate.Locked = True
  txtFirstName.Locked = True
  txtLastName.Locked = True
  txtPaymentType.Locked = True
  txtState.Locked = True
  txtZipCode.Locked = True
End Sub

Private Sub EnableTextBoxes()
  txtCustomerID.Locked = False
  txtAddress.Locked = False
  txtCity.Locked = False
  txtCreditCardNumber.Locked = False
  txtExpDate.Locked = False
  txtFirstName.Locked = False
  txtLastName.Locked = False
  txtPaymentType.Locked = False
  txtState.Locked = False
  txtZipCode.Locked = False
End Sub

Private Sub cmdAdd_Click()
    'Add a new record
  On Error GoTo HandleError
  If (cmdADD.Caption = "&Add") Then
    deABCStores.rsCustomers.AddNew
    DisableControls
    EnableTextBoxes
    setUpAdd
  Else
      'cancel the Add
    deABCStores.rsCustomers.CancelUpdate
    txtCustomerID.Locked = True
    EnableControls
    DisableTextBoxes
    cmdSave.Enabled = False
    cmdADD.Caption = "&Add"
    deABCStores.rsCustomers.MoveLast
    If deABCStores.rsCustomers.EOF Then
        deABCStores.rsCustomers.MovePrevious
    End If
  End If
cmdAdd_click_exit:
  Exit Sub
HandleError:
  MsgBox "Unable to carry out requested action.", _
          vbInformation, "ABC Stores"
End Sub

Private Sub cmdDelete_Click()
   'Delete the curent record
  Dim intYes As Single
  On Error GoTo HandleError
  intYes = MsgBox("WARNING RECORD DUE TO A PERMANENT DELETION: ARE YOU SURE ?", _
  vbYesNo, "ABC WARNING")
  If intYes = vbYes Then
    With deABCStores.rsCustomers
      .Delete
      .MoveNext
      If .EOF Then
        .MovePrevious
        If .BOF Then
          MsgBox "The recordset is empty.", _
            vbInformation, "No Records"
          DisableControls
        End If
      End If
    End With
  End If
cmddelete_click_exit:
  Exit Sub
HandleError:
  MsgBox "Unable to carry out requested action.", _
        vbInformation, "ABC Stores"
        On Error GoTo 0
End Sub

Private Sub cmdFirst_Click()
    'Move to the first record
  On Error Resume Next
  deABCStores.rsCustomers.MoveFirst
End Sub

Private Sub cmdLast_Click()
    'Move to the last record
  On Error Resume Next
  deABCStores.rsCustomers.MoveLast
End Sub

Private Sub cmdNext_Click()
    'Move to the next record
  On Error Resume Next
  With deABCStores.rsCustomers
    .MoveNext
    If (.EOF) Then
      .MoveLast
    End If
  End With
End Sub

Private Sub cmdPrevious_Click()
     'Move to the Previous record
  On Error Resume Next
  With deABCStores.rsCustomers
    .MovePrevious
    If (.BOF) Then
      .MoveFirst
    End If
  End With
End Sub

Private Sub cmdSave_Click()
        'Save the current record
    On Error GoTo HandleErrors
    deABCStores.rsCustomers.Update
    txtCustomerID.Locked = True
    EnableControls
    cmdSave.Enabled = False
    cmdADD.Caption = "&Add"
cmdsave_click_exit:
    Exit Sub
    
HandleErrors:
    Dim strMessage As String
    Dim errDBError As ADODB.Error
    For Each errDBError In deABCStores.conABCStores.Errors
        strMessage = strMessage & errDBError.Description & vbCrLf
    Next
    MsgBox strMessage, vbExclamation, "Duplicate Add"
        'Turn off error trapping
    On Error GoTo 0
End Sub

Private Sub Form_Activate()
  On Error Resume Next
  txtCustomerID.Locked = True
  With deABCStores.rsCustomers
      .MoveFirst
  End With
  cmdSave.Enabled = False
End Sub

Private Sub Form_GotFocus()
  On Error Resume Next
  txtCustomerID.Locked = True
  With deABCStores.rsCustomers
    .MoveFirst
  End With
  cmdSave.Enabled = False
End Sub

            'Data validation
Private Sub txtAddress_Validate(Cancel As Boolean)
    With txtAddress
        If .Text = "" Then
            MsgBox "Please Enter Address"
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
            Cancel = True
        End If
    End With
End Sub

Private Sub txtCity_Validate(Cancel As Boolean)
    With txtCity
        If .Text = "" Then
            MsgBox "Please Enter City"
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
            Cancel = True
        End If
    End With
End Sub

Private Sub txtCreditCardNumber_Validate(Cancel As Boolean)
    With txtCreditCardNumber
        If Not Len(.Text) = 16 Or Not IsNumeric(.Text) Then
            MsgBox "Please Enter 16 digits for Credit Card Number"
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
            Cancel = True
        End If
    End With
End Sub

Private Sub txtCustomerID_Validate(Cancel As Boolean)
    With txtCustomerID
        If Not Len(.Text) = 4 Or Not IsNumeric(.Text) Then
            MsgBox "Please Enter 4 digits for Customer ID"
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
            Cancel = True
        End If
    End With
End Sub

Private Sub txtExpDate_Validate(Cancel As Boolean)
      With txtExpDate
        If .Text = "" Then
            MsgBox "Please Enter Expiration Date"
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
            Cancel = True
        End If
    End With
End Sub

Private Sub txtFirstName_Validate(Cancel As Boolean)
    With txtFirstName
        If .Text = "" Then
            MsgBox "Please Enter First Name"
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
            Cancel = True
        End If
    End With
End Sub

Private Sub txtLastName_Validate(Cancel As Boolean)
    With txtLastName
        If .Text = "" Then
            MsgBox "Please Enter Last Name "
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
            Cancel = True
        End If
    End With
End Sub

Private Sub txtPaymentType_Validate(Cancel As Boolean)
    With txtPaymentType
        If .Text = "" Then
            MsgBox "Please Enter Payment Type"
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
            Cancel = True
        End If
    End With
End Sub

Private Sub txtState_Validate(Cancel As Boolean)
    With txtState
        If .Text = "" Then
            MsgBox "Please Enter State Abbreviation "
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
            Cancel = True
        End If
    End With
End Sub

Private Sub txtZipCode_Validate(Cancel As Boolean)
    With txtZipCode
        If Not Len(.Text) = 5 Or Not IsNumeric(.Text) Then
            MsgBox "Please Enter Zip Code "
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
            Cancel = True
        End If
    End With
End Sub
