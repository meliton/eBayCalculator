VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "eBay Calculator App 1.1 (5/31/2015)"
   ClientHeight    =   3480
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13170
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   13170
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmItmCst 
      BackColor       =   &H0080FF80&
      Caption         =   "Less Item Cost"
      Height          =   1335
      Left            =   5880
      TabIndex        =   28
      Top             =   1440
      Width           =   1935
      Begin VB.Label lblPercent1 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Caption         =   "lblPercent1"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label lblLessCost 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Caption         =   "lblLessCost"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   375
         Left            =   240
         TabIndex        =   29
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.Frame frmKillShipping 
      BackColor       =   &H0080FF80&
      Caption         =   "Profit - Less Fees n Shipping"
      Height          =   1335
      Left            =   2760
      TabIndex        =   25
      Top             =   1440
      Width           =   3015
      Begin VB.Label lblShip 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080FF80&
         Caption         =   "lblShip"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1920
         TabIndex        =   27
         Top             =   960
         Width           =   855
      End
      Begin VB.Label lblProfit 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Caption         =   "lblProfit"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   495
         Left            =   120
         TabIndex        =   26
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.Frame frmAfter 
      BackColor       =   &H0080FF80&
      Caption         =   "After eBay/PayPal Fees"
      Height          =   1335
      Left            =   120
      TabIndex        =   22
      Top             =   1440
      Width           =   2535
      Begin VB.Label lblPPeBayFees 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080FF80&
         Caption         =   "lblPPeBayFees"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   960
         TabIndex        =   24
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label lblAfterFees 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Caption         =   "lblAfterFees"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   495
         Left            =   240
         TabIndex        =   23
         Top             =   480
         Width           =   2055
      End
   End
   Begin VB.CommandButton btnTotal 
      Caption         =   "Get Totals"
      Height          =   495
      Left            =   8040
      TabIndex        =   21
      Top             =   2880
      Width           =   2535
   End
   Begin VB.Frame frmShipChrgd 
      BackColor       =   &H0080FF80&
      Caption         =   "Shipping Charged"
      Height          =   1095
      Left            =   5760
      TabIndex        =   17
      Top             =   120
      Width           =   2055
      Begin VB.TextBox txtShipCharged 
         Height          =   405
         Left            =   120
         TabIndex        =   2
         Text            =   "txtShipCharged"
         ToolTipText     =   "Shipping you are charging"
         Top             =   480
         Width           =   1815
      End
   End
   Begin VB.Frame frmSold 
      BackColor       =   &H0080FF80&
      Caption         =   "Will Sell/Sold Price"
      Height          =   1095
      Left            =   3480
      TabIndex        =   16
      Top             =   120
      Width           =   2175
      Begin VB.TextBox txtSell 
         Height          =   405
         Left            =   240
         TabIndex        =   1
         Text            =   "txtSell"
         ToolTipText     =   "Buy it Now or Auction Sold Price"
         Top             =   480
         Width           =   1815
      End
   End
   Begin VB.Frame frmItem 
      BackColor       =   &H0080FF80&
      Caption         =   "Item Cost"
      Height          =   1095
      Left            =   1920
      TabIndex        =   15
      Top             =   120
      Width           =   1455
      Begin VB.TextBox txtCost 
         Height          =   390
         Left            =   120
         TabIndex        =   0
         Text            =   "txtCost"
         ToolTipText     =   "Original cost you spent to purchase the item"
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Frame frmListing 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Listing Format"
      Height          =   1095
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   1695
      Begin VB.OptionButton optFixed 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Fixed Price"
         Height          =   285
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   1455
      End
      Begin VB.OptionButton optAuction 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Auction"
         Height          =   285
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame frmDiscounts 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Discounts"
      Height          =   1695
      Left            =   10680
      TabIndex        =   9
      ToolTipText     =   "Power Seller and Store type Discounts"
      Top             =   1680
      Width           =   2415
      Begin VB.ComboBox cboStore 
         Height          =   405
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   11
         ToolTipText     =   "Select Store Type"
         Top             =   840
         Width           =   2055
      End
      Begin VB.CheckBox chkPS 
         BackColor       =   &H00C0C0FF&
         Caption         =   "20% Power Seller"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         ToolTipText     =   "Power Seller Discount on Final Value Fees"
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label lblEbayFees 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0FF&
         Caption         =   "$0.00"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1440
         TabIndex        =   20
         Top             =   1320
         Width           =   855
      End
   End
   Begin VB.Frame frmShipping 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Actual Shipping Costs"
      Height          =   2655
      Left            =   8040
      TabIndex        =   7
      Top             =   120
      Width           =   2535
      Begin VB.TextBox txtShipping 
         Height          =   405
         Left            =   360
         TabIndex        =   4
         Text            =   "txtShipping"
         ToolTipText     =   "Actual Shipping Cost"
         Top             =   2040
         Width           =   1815
      End
      Begin VB.CheckBox chkUSPSDiscount 
         BackColor       =   &H00C0C0FF&
         Caption         =   "USPS Discount"
         Height          =   375
         Left            =   360
         TabIndex        =   8
         ToolTipText     =   "USPS Discount Prices"
         Top             =   480
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.ComboBox cboShipping 
         Height          =   405
         Left            =   360
         Style           =   2  'Dropdown List
         TabIndex        =   3
         ToolTipText     =   "Select USPS Shipping Cost"
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label lblOR 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0FF&
         Caption         =   "OR Enter Shipping"
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   1560
         Width           =   1935
      End
   End
   Begin VB.Frame frmPP 
      BackColor       =   &H00C0C0FF&
      Caption         =   "PayPal Transaction Fee"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   10680
      TabIndex        =   5
      Top             =   120
      Width           =   2415
      Begin VB.ComboBox cboPPFees 
         Height          =   405
         ItemData        =   "Form1.frx":0442
         Left            =   120
         List            =   "Form1.frx":0444
         Style           =   2  'Dropdown List
         TabIndex        =   6
         ToolTipText     =   "Select PayPal Transaction Fees"
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label lblPPFees 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0FF&
         Caption         =   "$0.00"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1440
         TabIndex        =   19
         Top             =   1080
         Width           =   855
      End
   End
   Begin VB.Label lblCalculator 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "lblCalculator"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   240
      TabIndex        =   31
      ToolTipText     =   "Click to open system calculator"
      Top             =   3000
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Program   : eBayCalculator.exe
' Date      : 1/2/2016
'
'
' Updated 5/31/2015 for new USPS postal rates
'
' TODO:  Use program name to set options, for example:
'        eBayCalculatorBP.exe could be Basic Store, Power Seller
'        eBayCalculatorN22.exe could be No Store, 2.2% PayPal transaction fee
'
' Listing Format       Discounts            PayPal Transaction Fee  USPS Discount
'  C -- auCtion        *P -- Power Seller   *29 -- 2.9% + $0.30     *U -- USPS Shipping Discount
' *F -- Fixed Price     N -- No Store        25 -- 2.5% + $0.30      X -- No Shipping Discount
'                      *B -- Basic Store     22 -- 2.2% + $0.30
'                       R -- pRemium Store   19 -- 1.9% + $0.30
'                       A -- Anchor Store    50 -- 5.0% + $0.05
' * - denotes default
'
'
' FIX: 20% Power Seller Discount doesn't work
'      $ - dollar signs disappear on all money calculations
'
'---------------------------------------------------------------------------------------

Option Explicit

Private Sub btnTotal_Click()
Call CheckBlankBoxes        'checks for blanks boxes and formats them
Call EbayPayPalTransFees    'calculate PayPal eBay transaction fees and shipping
End Sub

Private Sub cboShipping_GotFocus()
txtShipping.Text = "0.00"       'resets shipping box when USPS shipping is in focus
End Sub

Private Sub chkUSPSDiscount_Click()
Call Shipping
End Sub

Private Sub Form_Load()
Me.Height = 3870    'sets the form height
Call PayPalFees     'load PayPal Fees Combobox
Call Shipping       'load Shipping Combobox
Call StoreType      'load Store combobox
Call SetDefaults    'sets the default textboxes and radio settings
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set Form1 = Nothing
Unload Form1        'graceful exit
End Sub

Private Sub SetDefaults()
optFixed.Value = True           'set Fixed price radio button
txtCost.Text = "0.00"          'sets the cost value
txtSell.Text = "0.00"          'sets the sell value
txtShipCharged.Text = "0.00"   'sets the shipping charged value
txtShipping.Text = "0.00"      'sets the shipping box
lblCalculator.Caption = "Open CALCULATOR"   'sets calulator window
lblAfterFees.Caption = "$0.00"   'sets the after fees lbl
lblPPeBayFees.Caption = "$0.00"     'sets the PayPal/eBay total fees lbl
lblProfit.Caption = "$$$"          'sets profit label
lblShip.Caption = "$0.00"           'sets shipping lbl
lblLessCost.Caption = "-"           'sets less cost lbl
lblPercent1.Caption = "Profit / Item Cost"
End Sub

Private Sub PayPalFees()    'load PayPal Fees Combobox
cboPPFees.AddItem "2.9% + $0.30"
cboPPFees.AddItem "2.5% + $0.30"
cboPPFees.AddItem "2.2% + $0.30"
cboPPFees.AddItem "1.9% + $0.30"
cboPPFees.AddItem "5.0% + $0.30"
cboPPFees.ListIndex = 0     'picks the first item in the box
End Sub

Private Sub EbayPayPalTransFees()
On Error GoTo theEnd             'error handler
Dim intPercent As String
Dim intConst As String
Dim intNewTotal As String

intNewTotal = Val(txtSell.Text) + Val(txtShipCharged.Text)

intPercent = Left$(cboPPFees.Text, 3)
intConst = Right$(cboPPFees.Text, 3)
lblPPFees.Caption = Round(((intPercent / 100) * intNewTotal) + intConst, 2)

Dim intPercent1 As String
Dim intConst1 As String

intPercent1 = cboStore.Text  'gets the value of the Store
If intPercent1 = "No Store" Then
intConst1 = "0.1"
Else
intConst1 = "0.09"
End If

lblEbayFees.Caption = Round(intConst1 * intNewTotal, 2)
lblPPeBayFees = Val(lblEbayFees.Caption) + Val(lblPPFees.Caption)
lblAfterFees = intNewTotal - lblPPeBayFees.Caption
Call CheckShipping      'checks which shipping fee to use
lblProfit.Caption = Val(lblAfterFees.Caption) - Val(lblShip.Caption)
lblLessCost.Caption = Round(Val(lblProfit.Caption) - Val(txtCost.Text), 2)
Call PercentCalc

theEnd:     'basic error handler
End Sub

Private Sub CheckShipping()
On Error GoTo theEnd
If txtShipping.Text <> "0.00" Then
lblShip.Caption = txtShipping.Text
Else
lblShip.Caption = Right$(cboShipping.Text, 4)
End If

theEnd:     'basic error handler
End Sub

Private Sub PercentCalc()
On Error GoTo theEnd
If Val(txtCost.Text) > 0 Then
lblPercent1.Caption = Round(Val(lblProfit.Caption) / Val(txtCost.Text), 3)
Else
lblPercent1.Caption = "Profit / Item Cost"
End If

theEnd:     'basic error handler
End Sub

Private Sub CheckBlankBoxes()
If txtCost.Text = "" Then
txtCost.Text = "0.00"
End If

If txtSell.Text = "" Then
txtSell.Text = "0.00"
End If

If txtShipCharged.Text = "" Then
txtShipCharged.Text = "0.00"
End If

If txtShipping.Text = "" Then
txtShipping.Text = "0.00"
End If
End Sub

Private Sub Shipping()
txtShipping.Text = "0.00"           'resets the shipping box
If chkUSPSDiscount.Value = vbChecked Then   'its the cheaper postage
    cboShipping.Clear
    cboShipping.AddItem "1 oz, $2.04"
    cboShipping.AddItem "2 oz, $2.04"
    cboShipping.AddItem "3 oz, $2.04"
    cboShipping.AddItem "4 oz, $2.13"
    cboShipping.AddItem "5 oz, $2.22"
    cboShipping.AddItem "6 oz, $2.35"
    cboShipping.AddItem "7 oz, $2.53"
    cboShipping.AddItem "8 oz, $2.71"
    cboShipping.AddItem "9 oz, $2.89"
    cboShipping.AddItem "10 oz, $3.07"
    cboShipping.AddItem "11 oz, $3.25"
    cboShipping.AddItem "12 oz, $3.44"
    cboShipping.AddItem "13 oz, $3.63"
Else
    cboShipping.Clear       'its the expensive shipping
    cboShipping.AddItem "1 oz, $2.54"
    cboShipping.AddItem "2 oz, $2.54"
    cboShipping.AddItem "3 oz, $2.54"
    cboShipping.AddItem "4 oz, $2.74"
    cboShipping.AddItem "5 oz, $2.94"
    cboShipping.AddItem "6 oz, $3.14"
    cboShipping.AddItem "7 oz, $3.34"
    cboShipping.AddItem "8 oz, $3.54"
    cboShipping.AddItem "9 oz, $3.74"
    cboShipping.AddItem "10 oz, $3.94"
    cboShipping.AddItem "11 oz, $4.14"
    cboShipping.AddItem "12 oz, $4.34"
    cboShipping.AddItem "13 oz, $4.54"
End If
cboShipping.ListIndex = 0     'picks the first item in the box

End Sub

Private Sub StoreType()
cboStore.Clear              'clears the box before use
cboStore.AddItem "No Store"
cboStore.AddItem "Basic Store"
cboStore.AddItem "Premium Store"
cboStore.AddItem "Anchor Store"
cboStore.ListIndex = 1     'picks the second item in the box

End Sub

Private Sub lblCalculator_Click()
On Error Resume Next        'if it doesn't open, just continue
Shell "calc", vbNormalFocus
End Sub

Private Sub txtCost_DblClick()
txtCost.Text = ""   'clears the box
End Sub

Private Sub txtSell_DblClick()
txtSell.Text = ""   'clears the box
End Sub

Private Sub txtShipCharged_DblClick()
txtShipCharged.Text = ""    'clears the box
End Sub

Private Sub txtShipping_DblClick()
txtShipping.Text = ""   'clears the box
End Sub

Private Sub txtCost_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then       'if Enter is pressed, go to the next box and clear it
    txtSell.Text = ""           'clears the txtSell box
    txtSell.SetFocus
    End If
    If KeyAscii < 48 Or KeyAscii > 57 Then  'allows 0-9
      If KeyAscii <> 46 Then                'allows the decimal
      If KeyAscii <> 8 Then                 'allow the backspace
       KeyAscii = 0
      End If
    End If
    End If
End Sub

Private Sub txtSell_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then       'if Enter is pressed, go to the next box
    txtShipCharged.SetFocus
    End If
    If KeyAscii < 48 Or KeyAscii > 57 Then  'allows 0-9
      If KeyAscii <> 46 Then                'allows the decimal
      If KeyAscii <> 8 Then                 'allow the backspace
       KeyAscii = 0
      End If
    End If
    End If
End Sub

Private Sub txtShipCharged_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then       'if Enter is pressed, go to the next box
    btnTotal.SetFocus           'sets focus to total button
    End If
    If KeyAscii < 48 Or KeyAscii > 57 Then  'allows 0-9
      If KeyAscii <> 46 Then                'allows the decimal
      If KeyAscii <> 8 Then                 'allow the backspace
       KeyAscii = 0
      End If
    End If
    End If
End Sub

Private Sub txtShipping_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then       'if Enter is pressed, go to the next box
    btnTotal.SetFocus           'sets focus to total button
    End If
    If KeyAscii < 48 Or KeyAscii > 57 Then  'allows 0-9
      If KeyAscii <> 46 Then                'allows the decimal
      If KeyAscii <> 8 Then                 'allow the backspace
       KeyAscii = 0
      End If
    End If
    End If
End Sub
