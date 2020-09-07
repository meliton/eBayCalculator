VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "eBay Calculator App 1.6 (1/26/2020 Rates)"
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12585
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
   ScaleHeight     =   3765
   ScaleWidth      =   12585
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmItmCst 
      BackColor       =   &H0080FF80&
      Caption         =   "Less Item Cost"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   5640
      TabIndex        =   24
      Top             =   1320
      Width           =   1695
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
         TabIndex        =   26
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label lblLessCost 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Caption         =   "lblLessCost"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   495
         Left            =   120
         TabIndex        =   25
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame frmKillShipping 
      BackColor       =   &H0080FF80&
      Caption         =   "Profit - Less Fees n Shipping"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   2640
      TabIndex        =   21
      Top             =   1320
      Width           =   2895
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
         Left            =   1200
         TabIndex        =   23
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label lblProfit 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Caption         =   "lblProfit"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   615
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.Frame frmAfter 
      BackColor       =   &H0080FF80&
      Caption         =   "After eBay/PayPal Fees"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   18
      Top             =   1320
      Width           =   2415
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
         Left            =   480
         TabIndex        =   20
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label lblAfterFees 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Caption         =   "lblAfterFees"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   495
         Left            =   240
         TabIndex        =   19
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.CommandButton btnTotal 
      Caption         =   "Get Totals"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4680
      TabIndex        =   17
      Top             =   2880
      Width           =   2655
   End
   Begin VB.Frame frmShipChrgd 
      BackColor       =   &H0080FF80&
      Caption         =   "Shipping Charged"
      Height          =   1095
      Left            =   4800
      TabIndex        =   13
      Top             =   120
      Width           =   2535
      Begin VB.TextBox txtShipCharged 
         Height          =   405
         Left            =   120
         TabIndex        =   2
         Text            =   "txtShipCharged"
         ToolTipText     =   "Shipping you are charging"
         Top             =   480
         Width           =   2055
      End
   End
   Begin VB.Frame frmSold 
      BackColor       =   &H0080FF80&
      Caption         =   "Will Sell/Sold Price"
      Height          =   1095
      Left            =   2280
      TabIndex        =   12
      Top             =   120
      Width           =   2415
      Begin VB.TextBox txtSell 
         Height          =   405
         Left            =   240
         TabIndex        =   1
         Text            =   "txtSell"
         ToolTipText     =   "Buy it Now or Auction Sold Price"
         Top             =   480
         Width           =   1935
      End
   End
   Begin VB.Frame frmItem 
      BackColor       =   &H0080FF80&
      Caption         =   "Item Cost"
      Height          =   1095
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   2055
      Begin VB.TextBox txtCost 
         Height          =   390
         Left            =   120
         TabIndex        =   0
         Text            =   "txtCost"
         ToolTipText     =   "Original cost you spent to purchase the item"
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.Frame frmDiscounts 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Discounts"
      Height          =   2055
      Left            =   10080
      TabIndex        =   9
      ToolTipText     =   "Power Seller and Store type Discounts"
      Top             =   1560
      Width           =   2415
      Begin VB.CheckBox chkStore 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Has Store?"
         Height          =   285
         Left            =   120
         TabIndex        =   29
         ToolTipText     =   "Check for Basic, Premium, and Anchor Store "
         Top             =   840
         Width           =   1575
      End
      Begin VB.ComboBox cboFVF 
         Height          =   405
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   28
         ToolTipText     =   "Select Final Value Fee"
         Top             =   1320
         Width           =   1455
      End
      Begin VB.CheckBox chkPS 
         BackColor       =   &H00C0C0FF&
         Caption         =   "10% Power Seller"
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
         TabIndex        =   16
         Top             =   1440
         Width           =   855
      End
   End
   Begin VB.Frame frmShipping 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Actual Shipping Costs"
      Height          =   3495
      Left            =   7440
      TabIndex        =   7
      Top             =   120
      Width           =   2535
      Begin VB.OptionButton optZ8 
         BackColor       =   &H00C0C0FF&
         Caption         =   "8"
         Height          =   495
         Left            =   1560
         TabIndex        =   36
         ToolTipText     =   "Zone 8 & 9"
         Top             =   1800
         Width           =   255
      End
      Begin VB.OptionButton optZ7 
         BackColor       =   &H00C0C0FF&
         Caption         =   "7"
         Height          =   495
         Left            =   1320
         TabIndex        =   35
         ToolTipText     =   "Zone 7"
         Top             =   1800
         Width           =   255
      End
      Begin VB.OptionButton optZ6 
         BackColor       =   &H00C0C0FF&
         Caption         =   "6"
         Height          =   495
         Left            =   1080
         TabIndex        =   34
         ToolTipText     =   "Zone 6"
         Top             =   1800
         Width           =   255
      End
      Begin VB.OptionButton optZ5 
         BackColor       =   &H00C0C0FF&
         Caption         =   "5"
         Height          =   495
         Left            =   840
         TabIndex        =   33
         ToolTipText     =   "Zone 5"
         Top             =   1800
         Width           =   255
      End
      Begin VB.OptionButton optZ4 
         BackColor       =   &H00C0C0FF&
         Caption         =   "4"
         Height          =   495
         Left            =   600
         TabIndex        =   32
         ToolTipText     =   "Zone 4"
         Top             =   1800
         Width           =   255
      End
      Begin VB.OptionButton optZ3 
         BackColor       =   &H00C0C0FF&
         Caption         =   "3"
         Height          =   495
         Left            =   360
         TabIndex        =   31
         ToolTipText     =   "Zone 3"
         Top             =   1800
         Width           =   255
      End
      Begin VB.OptionButton optZ1 
         BackColor       =   &H00C0C0FF&
         Caption         =   "1"
         Height          =   495
         Left            =   120
         TabIndex        =   30
         ToolTipText     =   "Zone 1 & 2"
         Top             =   1800
         Width           =   255
      End
      Begin VB.TextBox txtShipping 
         Height          =   405
         Left            =   120
         TabIndex        =   4
         Text            =   "txtShipping"
         ToolTipText     =   "Actual Shipping Cost"
         Top             =   3000
         Width           =   1815
      End
      Begin VB.CheckBox chkUSPSDiscount 
         BackColor       =   &H00C0C0FF&
         Caption         =   "USPS Discount"
         Height          =   375
         Left            =   360
         TabIndex        =   8
         ToolTipText     =   "USPS Discount Prices"
         Top             =   360
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.ComboBox cboShipping 
         Height          =   405
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   3
         ToolTipText     =   "Select USPS Shipping Cost"
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label lblOR 
         BackColor       =   &H00C0C0FF&
         Caption         =   "OR Enter Shipping"
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   2640
         Width           =   2175
      End
      Begin VB.Label lblZone 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Select Zone Below"
         Height          =   375
         Left            =   120
         TabIndex        =   37
         Top             =   1560
         Width           =   1815
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
      Height          =   1335
      Left            =   10080
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
         Top             =   360
         Width           =   1815
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
         TabIndex        =   15
         Top             =   840
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
      TabIndex        =   27
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
' Date      : 9/07/2020
'
' Updated 5/31/2015 for new USPS postal rates
' Updated 1/17/2016 for new USPS postal rates
' Updated 2/10/2016 for new USPS postal rates
' Updates 4/24/2017 for new USPS postal rates
' Updates 5/01/2017 for new eBay Fees, also added Flat Rate USPS prices
' Updates 1/21/2018 for new USPS postal rates
' Updates 9/07/2020 for new USPS postal rates and remove Listing Format tile
'
' TODO:  Use program name to set options, for example:
'        eBayCalculatorSP.exe could be Basic Store, Power Seller
'        eBayCalculatorN22.exe could be No Store, 2.2% PayPal transaction fee
'
' USPS Discount         Discounts            PayPal Transaction Fee  Final Value Fees
' *U -- USPS Discount  *P -- Power Seller   *29 -- 2.9% + $0.30     1.5  -- 1.5%  FV Fee
'  X -- No Discount     N -- No Store        25 -- 2.5% + $0.30     3.5  -- 3.5%  FV Fee
'                      *S -- Store           22 -- 2.2% + $0.30     4.0  -- 4.0%  FV Fee
'                                            19 -- 1.9% + $0.30     6.15 -- 6.15% FV Fee
'                                            50 -- 5.0% + $0.05    *8.15 -- 8.15% FV Fee
' * - denotes default                                               9.15 -- 9.15% FV Fee
'                                                                   10 - 10% FV Fee (No Store ONLY)
'
' FIX: Copy/Paste Text into Number Fields, bug
'      Dollar decimals short/long, bug
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

Private Sub chkStore_Click()
If chkStore = 1 Then    'Has Store is checked
cboFVF.Visible = True
Else
cboFVF.Visible = False
End If
End Sub

Private Sub chkUSPSDiscount_Click()
Call Shipping
End Sub

Private Sub Form_Load()
Me.Height = 4230    'sets the form height
Call PayPalFees     'load PayPal Fees Combobox
Call Shipping       'load Shipping Combobox
Call FinalValueFees 'load Final Value Fees percentage Combobox
Call SetDefaults    'sets the default textboxes and radio settings
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set Form1 = Nothing
Unload Form1        'graceful exit
End Sub

Private Sub SetDefaults()
chkStore.Value = 1             'store is checked
chkPS.Value = 1                'PowerSeller is checked
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
optZ1 = True                    'sets shipping Zone 1
chkUSPSDiscount = 1             'sets USPS discount box

End Sub

Private Sub PayPalFees()    'load PayPal Fees Combobox
cboPPFees.AddItem "2.9% + $0.30"
cboPPFees.AddItem "2.5% + $0.30"
cboPPFees.AddItem "2.2% + $0.30"
cboPPFees.AddItem "1.9% + $0.30"
cboPPFees.AddItem "5.0% + $0.05"
cboPPFees.ListIndex = 0     'picks the first item in the box
End Sub

Private Sub FinalValueFees()  'load Final Value Fees Combobox
cboFVF.AddItem "1.50% FVF"
cboFVF.AddItem "3.50% FVF"
cboFVF.AddItem "4.00% FVF"
cboFVF.AddItem "6.15% FVF"
cboFVF.AddItem "8.15% FVF"
cboFVF.AddItem "9.15% FVF"
cboFVF.ListIndex = 5   'picks the last item in the box
End Sub

Private Sub EbayPayPalTransFees()
On Error GoTo theEnd             'error handler
Dim intPercent As String
Dim intConst As String
Dim intNewTotal As String
Dim intFVF As String
Dim intPowerSeller As String

intNewTotal = Val(txtSell.Text) + Val(txtShipCharged.Text)

intFVF = Left$(cboFVF.Text, 4)
intPercent = Left$(cboPPFees.Text, 3)
intConst = Right$(cboPPFees.Text, 3)
lblPPFees.Caption = Round(((intPercent / 100) * intNewTotal) + intConst, 2)

Dim intPercent1 As String
Dim intConst1 As String

intPercent1 = chkStore.Value  'gets the value of the Store
If intPercent1 = 0 Then       'Has Store?, not checked (No Store)
intConst1 = "0.1"             'No Store fee
Else
intConst1 = 0.01 * Val(intFVF)    'Store discount fee
End If

If chkPS.Value = 1 Then      'PowerSeller is checked
intPowerSeller = "0.9"       '10% PowerSeller discount (May 1, 2017)
Else
intPowerSeller = "1"         'no PowerSeller discount
End If

lblEbayFees.Caption = Round(intPowerSeller * intConst1 * intNewTotal, 2)
lblPPeBayFees = Val(lblEbayFees.Caption) + Val(lblPPFees.Caption)
lblAfterFees = intNewTotal - lblPPeBayFees.Caption
Call CheckShipping      'checks which shipping fee to use
lblProfit.Caption = Val(lblAfterFees.Caption) - Val(lblShip.Caption)
lblLessCost.Caption = Round(Val(lblProfit.Caption) - Val(txtCost.Text), 2)
Call PercentCalc

'Add $$$ here
lblAfterFees = "$" & lblAfterFees.Caption
lblProfit = "$" & lblProfit.Caption
lblLessCost = "$" & lblLessCost.Caption
lblPPeBayFees = "-$" & lblPPeBayFees.Caption
lblShip = "-$" & lblShip.Caption
lblPPFees = "-$" & lblPPFees.Caption
lblEbayFees = "-$" & lblEbayFees.Caption


theEnd:     'basic error handler
End Sub

Private Sub CheckShipping()
On Error GoTo theEnd
If txtShipping.Text <> "0.00" Then
lblShip.Caption = txtShipping.Text
Else
Dim intStrLen As Integer        ' length of string in combobox
Dim intDollarSign As Integer    ' location of dollar sign in combobox

intStrLen = Len(cboShipping.Text)
intDollarSign = InStr(1, cboShipping.Text, "$")
lblShip.Caption = Mid(cboShipping.Text, intDollarSign + 1, intStrLen)
End If

theEnd:     'basic error handler
End Sub

Private Sub PercentCalc()
On Error GoTo theEnd
If Val(txtCost.Text) > 0 Then
lblPercent1.Caption = Round(Val(lblProfit.Caption) / Val(txtCost.Text), 3)
lblPercent1 = lblPercent1.Caption & " times"
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
If chkUSPSDiscount.Value = vbChecked And optZ1 = True Then   'its the cheaper postage
    cboShipping.Clear
    cboShipping.AddItem "1 - 4 oz, $2.74"
    cboShipping.AddItem "5 - 8 oz, $3.21"
    cboShipping.AddItem "9 - 12 oz, $3.93"
    cboShipping.AddItem "13 - 16 oz, $5.04"
ElseIf chkUSPSDiscount.Value = vbChecked And optZ3 = True Then
    cboShipping.Clear
    cboShipping.AddItem "1 - 4 oz, $2.76"
    cboShipping.AddItem "5 - 8 oz, $3.23"
    cboShipping.AddItem "9 - 12 oz, $3.97"
    cboShipping.AddItem "13 - 16 oz, $5.08"
ElseIf chkUSPSDiscount.Value = vbChecked And optZ4 = True Then
    cboShipping.Clear
    cboShipping.AddItem "1 - 4 oz, $2.78"
    cboShipping.AddItem "5 - 8 oz, $3.25"
    cboShipping.AddItem "9 - 12 oz, $4.00"
    cboShipping.AddItem "13 - 16 oz, $5.12"
ElseIf chkUSPSDiscount.Value = vbChecked And optZ5 = True Then
    cboShipping.Clear
    cboShipping.AddItem "1 - 4 oz, $2.84"
    cboShipping.AddItem "5 - 8 oz, $3.31"
    cboShipping.AddItem "9 - 12 oz, $4.08"
    cboShipping.AddItem "13 - 16 oz, $5.27"
ElseIf chkUSPSDiscount.Value = vbChecked And optZ6 = True Then
    cboShipping.Clear
    cboShipping.AddItem "1 - 4 oz, $2.93"
    cboShipping.AddItem "5 - 8 oz, $3.39"
    cboShipping.AddItem "9 - 12 oz, $4.18"
    cboShipping.AddItem "13 - 16 oz, $5.40"
ElseIf chkUSPSDiscount.Value = vbChecked And optZ7 = True Then
    cboShipping.Clear
    cboShipping.AddItem "1 - 4 oz, $3.05"
    cboShipping.AddItem "5 - 8 oz, $3.52"
    cboShipping.AddItem "9 - 12 oz, $4.32"
    cboShipping.AddItem "13 - 16 oz, $5.54"
ElseIf chkUSPSDiscount.Value = vbChecked And optZ8 = True Then
    cboShipping.Clear
    cboShipping.AddItem "1 - 4 oz, $3.18"
    cboShipping.AddItem "5 - 8 oz, $3.67"
    cboShipping.AddItem "9 - 12 oz, $4.46"
    cboShipping.AddItem "13 - 16 oz, $5.70"
ElseIf chkUSPSDiscount.Value = vbUnchecked And optZ1 = True Then 'more expensive shipping
    cboShipping.Clear
    cboShipping.AddItem "1 - 4 oz, $3.80"
    cboShipping.AddItem "5 - 8 oz, $4.60"
    cboShipping.AddItem "9 - 12 oz, $5.30"
    cboShipping.AddItem "13 oz, $5.90"
ElseIf chkUSPSDiscount.Value = vbUnchecked And optZ3 = True Then
    cboShipping.Clear
    cboShipping.AddItem "1 - 4 oz, $3.85"
    cboShipping.AddItem "5 - 8 oz, $4.65"
    cboShipping.AddItem "9 - 12 oz, $5.35"
    cboShipping.AddItem "13 oz, $5.95"
ElseIf chkUSPSDiscount.Value = vbUnchecked And optZ4 = True Then
    cboShipping.Clear
    cboShipping.AddItem "1 - 4 oz, $3.90"
    cboShipping.AddItem "5 - 8 oz, $4.70"
    cboShipping.AddItem "9 - 12 oz, $5.40"
    cboShipping.AddItem "13 oz, $6.05"
ElseIf chkUSPSDiscount.Value = vbUnchecked And optZ5 = True Then
    cboShipping.Clear
    cboShipping.AddItem "1 - 4 oz, $3.95"
    cboShipping.AddItem "5 - 8 oz, $4.75"
    cboShipping.AddItem "9 - 12 oz, $5.45"
    cboShipping.AddItem "13 oz, $6.15"
ElseIf chkUSPSDiscount.Value = vbUnchecked And optZ6 = True Then
    cboShipping.Clear
    cboShipping.AddItem "1 - 4 oz, $4.00"
    cboShipping.AddItem "5 - 8 oz, $4.80"
    cboShipping.AddItem "9 - 12 oz, $5.50"
    cboShipping.AddItem "13 oz, $6.20"
ElseIf chkUSPSDiscount.Value = vbUnchecked And optZ7 = True Then
    cboShipping.Clear
    cboShipping.AddItem "1 - 4 oz, $4.05"
    cboShipping.AddItem "5 - 8 oz, $4.90"
    cboShipping.AddItem "9 - 12 oz, $5.65"
    cboShipping.AddItem "13 oz, $6.40"
ElseIf chkUSPSDiscount.Value = vbUnchecked And optZ8 = True Then
    cboShipping.Clear
    cboShipping.AddItem "1 - 4 oz, $4.20"
    cboShipping.AddItem "5 - 8 oz, $5.00"
    cboShipping.AddItem "9 - 12 oz, $5.75"
    cboShipping.AddItem "13 oz, $6.50"
End If
    cboShipping.AddItem "Envelope, $7.15"
    cboShipping.AddItem "Small Box, $7.65"
    cboShipping.AddItem "Med. Box, $13.25"
    cboShipping.AddItem "Large Box, $18.30"
    cboShipping.ListIndex = 0     'picks the first item in the box

End Sub

Private Sub lblCalculator_Click()
On Error Resume Next        'if it doesn't open, just continue
Shell "calc", vbNormalFocus
End Sub

Private Sub optZ1_Click()
Call Shipping
End Sub

Private Sub optZ3_Click()
Call Shipping
End Sub

Private Sub optZ4_Click()
Call Shipping
End Sub

Private Sub optZ5_Click()
Call Shipping
End Sub

Private Sub optZ6_Click()
Call Shipping
End Sub

Private Sub optZ7_Click()
Call Shipping
End Sub

Private Sub optZ8_Click()
Call Shipping
End Sub

Private Sub optZ9_Click()
Call Shipping
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
