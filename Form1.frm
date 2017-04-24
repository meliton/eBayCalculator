VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "eBay Calculator App 1.3 (2/10/2016)"
   ClientHeight    =   3480
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13050
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
   ScaleWidth      =   13050
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmItmCst 
      BackColor       =   &H0080FF80&
      Caption         =   "Less Item Cost"
      Height          =   1335
      Left            =   5880
      TabIndex        =   27
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
         TabIndex        =   29
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
         TabIndex        =   28
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.Frame frmKillShipping 
      BackColor       =   &H0080FF80&
      Caption         =   "Profit - Less Fees n Shipping"
      Height          =   1335
      Left            =   2760
      TabIndex        =   24
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
         Left            =   1440
         TabIndex        =   26
         Top             =   960
         Width           =   1335
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
         TabIndex        =   25
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.Frame frmAfter 
      BackColor       =   &H0080FF80&
      Caption         =   "After eBay/PayPal Fees"
      Height          =   1335
      Left            =   120
      TabIndex        =   21
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
         TabIndex        =   23
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
         TabIndex        =   22
         Top             =   480
         Width           =   2055
      End
   End
   Begin VB.CommandButton btnTotal 
      Caption         =   "Get Totals"
      Height          =   495
      Left            =   7920
      TabIndex        =   20
      Top             =   2880
      Width           =   2535
   End
   Begin VB.Frame frmShipChrgd 
      BackColor       =   &H0080FF80&
      Caption         =   "Shipping Charged"
      Height          =   1095
      Left            =   5760
      TabIndex        =   16
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
      TabIndex        =   15
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
      TabIndex        =   14
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
      TabIndex        =   11
      Top             =   120
      Width           =   1695
      Begin VB.OptionButton optFixed 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Fixed Price"
         Height          =   285
         Left            =   120
         TabIndex        =   13
         Top             =   720
         Width           =   1455
      End
      Begin VB.OptionButton optAuction 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Auction"
         Height          =   285
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame frmDiscounts 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Discounts"
      Height          =   1935
      Left            =   10560
      TabIndex        =   9
      ToolTipText     =   "Power Seller and Store type Discounts"
      Top             =   1440
      Width           =   2415
      Begin VB.CheckBox chkStore 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Has Store?"
         Height          =   285
         Left            =   120
         TabIndex        =   32
         ToolTipText     =   "Check for Basic, Premium, and Anchor Store "
         Top             =   840
         Width           =   1575
      End
      Begin VB.ComboBox cboFVF 
         Height          =   405
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   31
         ToolTipText     =   "Select Final Value Fee"
         Top             =   1320
         Width           =   1215
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
         TabIndex        =   19
         Top             =   1440
         Width           =   855
      End
   End
   Begin VB.Frame frmShipping 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Actual Shipping Costs"
      Height          =   2655
      Left            =   7920
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
         TabIndex        =   17
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
      Height          =   1215
      Left            =   10560
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
         TabIndex        =   18
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
      TabIndex        =   30
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
' Updated 1/17/2016 for new USPS postal rates
' Updated 2/10/2016 for new USPS postal rates
'
' TODO:  Use program name to set options, for example:
'        eBayCalculatorSP.exe could be Basic Store, Power Seller
'        eBayCalculatorN22.exe could be No Store, 2.2% PayPal transaction fee
'
' Listing Format        Discounts            PayPal Transaction Fee  Final Value Fees
'  A -- Auction         *P -- Power Seller   *29 -- 2.9% + $0.30     4 -- 4% FV Fee
' *F -- Fixed Price      N -- No Store        25 -- 2.5% + $0.30     6 -- 6% FV Fee
'                       *S -- Store           22 -- 2.2% + $0.30     7 -- 7% FV Fee
' USPS Discount                               19 -- 1.9% + $0.30     8 -- 8% FV Fee
' *U -- USPS Discount                         50 -- 5.0% + $0.05    *9 -- 9% FV Fee
'  X -- No Discount                                                 10 - 10% FV Fee (No Store ONLY)
'
' * - denotes default
'
'
' FIX: $ - dollar signs disappear on all money calculations, bug
'      Copy/Paste Text into Number Fields, bug
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
Me.Height = 3900    'sets the form height
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
optFixed.Value = True          'set Fixed price radio button
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
cboFVF.AddItem "4% FVF"
cboFVF.AddItem "6% FVF"
cboFVF.AddItem "7% FVF"
cboFVF.AddItem "8% FVF"
cboFVF.AddItem "9% FVF"
cboFVF.ListIndex = 4   'picks the last item in the box
End Sub

Private Sub EbayPayPalTransFees()
On Error GoTo theEnd             'error handler
Dim intPercent As String
Dim intConst As String
Dim intNewTotal As String
Dim intFVF As String
Dim intPowerSeller As String

intNewTotal = Val(txtSell.Text) + Val(txtShipCharged.Text)

intFVF = Left$(cboFVF.Text, 1)
intPercent = Left$(cboPPFees.Text, 3)
intConst = Right$(cboPPFees.Text, 3)
lblPPFees.Caption = Round(((intPercent / 100) * intNewTotal) + intConst, 2)

Dim intPercent1 As String
Dim intConst1 As String

intPercent1 = chkStore.Value  'gets the value of the Store
If intPercent1 = 0 Then       'Has Store?, not checked (No Store)
intConst1 = "0.1"             'No Store fee
Else
intConst1 = "0.0" & intFVF    'Store discount fee
End If

If chkPS.Value = 1 Then      'PowerSeller is checked
intPowerSeller = "0.8"       '20% PowerSeller discount
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
    cboShipping.AddItem "1 oz, $2.54"
    cboShipping.AddItem "2 oz, $2.54"
    cboShipping.AddItem "3 oz, $2.54"
    cboShipping.AddItem "4 oz, $2.60"
    cboShipping.AddItem "5 oz, $2.60"
    cboShipping.AddItem "6 oz, $2.60"
    cboShipping.AddItem "7 oz, $2.60"
    cboShipping.AddItem "8 oz, $2.60"
    cboShipping.AddItem "9 oz, $3.30"
    cboShipping.AddItem "10 oz, $3.35"
    cboShipping.AddItem "11 oz, $3.40"
    cboShipping.AddItem "12 oz, $3.45"
    cboShipping.AddItem "13 oz, $3.50"
Else
    cboShipping.Clear       'its the expensive shipping
    cboShipping.AddItem "1 oz, $2.67"
    cboShipping.AddItem "2 oz, $2.67"
    cboShipping.AddItem "3 oz, $2.67"
    cboShipping.AddItem "4 oz, $2.67"
    cboShipping.AddItem "5 oz, $2.85"
    cboShipping.AddItem "6 oz, $3.03"
    cboShipping.AddItem "7 oz, $3.21"
    cboShipping.AddItem "8 oz, $3.39"
    cboShipping.AddItem "9 oz, $3.57"
    cboShipping.AddItem "10 oz, $3.75"
    cboShipping.AddItem "11 oz, $3.93"
    cboShipping.AddItem "12 oz, $4.11"
    cboShipping.AddItem "13 oz, $4.29"
End If
cboShipping.ListIndex = 0     'picks the first item in the box

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
