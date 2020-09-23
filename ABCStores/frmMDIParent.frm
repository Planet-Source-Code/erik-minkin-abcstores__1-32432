VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMDIParent 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "ABC Stores"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   4680
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.Toolbar tlbABC 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ilsToolBar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "back"
            Object.ToolTipText     =   "Activate Customer Form"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "forward"
            Object.ToolTipText     =   "Activate Products form"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "report"
            Object.ToolTipText     =   "Report"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "closeActive"
            Object.ToolTipText     =   "Close Active Form"
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsToolBar 
      Left            =   2040
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIParent.frx":0000
            Key             =   "back"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIParent.frx":0114
            Key             =   "forward"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIParent.frx":0228
            Key             =   "report"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIParent.frx":0348
            Key             =   "closeActive"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuForm 
      Caption         =   "F&orm"
      Begin VB.Menu mnuFormCustomers 
         Caption         =   "&Customers"
      End
      Begin VB.Menu mnuFormProducts 
         Caption         =   "P&roducts"
      End
      Begin VB.Menu mnuFormCloseCurrentForm 
         Caption         =   "&Close Current Form"
      End
   End
   Begin VB.Menu mnuReport 
      Caption         =   "&Report"
      Begin VB.Menu mnuReportCustomers 
         Caption         =   "&Customers"
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnuWindowCascade 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu mnuWindowTileHorizontally 
         Caption         =   "Tile &Horizontally"
      End
      Begin VB.Menu mnuWindowTileVertically 
         Caption         =   "Tile &Vertically"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMDIParent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project ABCStores
'Form MDIParent
'Author Erik Minkin

Option Explicit

Private Sub mnuFileExit_Click()
  Dim aForm As Form
  For Each aForm In Forms
    Unload aForm
  Next
End Sub

Private Sub mnuFormCloseCurrentForm_Click()
  On Error GoTo HandleError
  Unload ActiveForm
  Exit Sub
HandleError:
    MsgBox "There is no active form", vbInformation, "Action"
    Resume Next
End Sub

Private Sub mnuFormCustomers_Click()
  frmCustomers.Show
  frmCustomers.SetFocus
End Sub

Private Sub mnuFormProducts_Click()
  frmProducts.Show
  frmProducts.SetFocus
End Sub

Private Sub mnuHelpAbout_Click()
  frmAbout.Show
  frmAbout.SetFocus
End Sub

Private Sub mnuReportCustomers_Click()
    frmCustomers.Hide
    rptReport.Show
End Sub

Private Sub mnuWindowCascade_Click()
  frmMDIParent.Arrange vbCascade
End Sub

Private Sub mnuWindowTileHorizontally_Click()
  frmMDIParent.Arrange vbHorizontal
End Sub

Private Sub mnuWindowTileVertically_Click()
  frmMDIParent.Arrange vbVertical
End Sub

Private Sub tlbABC_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "back"
            Call mnuFormCustomers_Click
        Case "forward"
            Call mnuFormProducts_Click
        Case "report"
            Call mnuReportCustomers_Click
        Case "closeActive"
            Call mnuFormCloseCurrentForm_Click
    End Select
End Sub
