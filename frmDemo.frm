VERSION 5.00
Begin VB.Form frmDemo 
   Caption         =   "Demonstraton of the PostalAddress Control"
   ClientHeight    =   3795
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6090
   LinkTopic       =   "Form1"
   ScaleHeight     =   3795
   ScaleWidth      =   6090
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtStateName 
      BackColor       =   &H80000016&
      Height          =   315
      Left            =   1740
      Locked          =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1260
      Width           =   4155
   End
   Begin VB.TextBox txtErrorDescription 
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   1740
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1620
      Width           =   4155
   End
   Begin VB.TextBox txtZip4 
      Height          =   315
      Left            =   5400
      TabIndex        =   11
      Top             =   3060
      Width           =   495
   End
   Begin VB.TextBox txtZip 
      Height          =   315
      Left            =   4740
      TabIndex        =   10
      Top             =   3060
      Width           =   555
   End
   Begin VB.TextBox txtState 
      Height          =   315
      Left            =   4020
      TabIndex        =   7
      Top             =   3060
      Width           =   495
   End
   Begin VB.TextBox txtCity 
      Height          =   315
      Left            =   1740
      TabIndex        =   5
      Top             =   3060
      Width           =   2175
   End
   Begin VB.TextBox txtStreetLine2 
      Height          =   315
      Left            =   1740
      TabIndex        =   4
      Top             =   2700
      Width           =   4155
   End
   Begin VB.TextBox txtStreetLine1 
      Height          =   315
      Left            =   1740
      TabIndex        =   3
      Top             =   2340
      Width           =   4155
   End
   Begin Demo.PostalAddress ctlPostalAddress 
      Height          =   675
      Left            =   1740
      TabIndex        =   0
      Top             =   180
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   1191
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblGeneral 
      Caption         =   "State Name:"
      Height          =   255
      Index           =   6
      Left            =   240
      TabIndex        =   15
      Top             =   1320
      Width           =   1035
   End
   Begin VB.Label lblGeneral 
      Caption         =   "Error Description:"
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   13
      Top             =   1680
      Width           =   1275
   End
   Begin VB.Line Line1 
      X1              =   180
      X2              =   5940
      Y1              =   2100
      Y2              =   2100
   End
   Begin VB.Label lblGeneral 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   5340
      TabIndex        =   9
      Top             =   3120
      Width           =   195
   End
   Begin VB.Label lblGeneral 
      Caption         =   ","
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   3960
      TabIndex        =   8
      Top             =   3120
      Width           =   195
   End
   Begin VB.Label lblGeneral 
      Caption         =   "City, State ZIP:"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   6
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label lblGeneral 
      Caption         =   "Street Address:"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   2
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label lblGeneral 
      Caption         =   "Address:          (This is the control)"
      Height          =   495
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "frmDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ctlPostalAddress_Change()
    If Me.ActiveControl Is ctlPostalAddress Then AddressToSeparateFields
    
    If ctlPostalAddress.IsValid Then
        txtErrorDescription.BackColor = vbButtonFace
    Else
        txtErrorDescription.BackColor = vbInfoBackground
    End If
    txtErrorDescription.Text = ctlPostalAddress.ValidationError
    txtStateName.Text = ctlPostalAddress.StateName(ctlPostalAddress.State)
End Sub

Private Sub Form_Load()
    ctlPostalAddress.Address = "7256 Whitney Way" & vbCrLf _
      & "Madison, WI 53562"
    AddressToSeparateFields
End Sub

Private Sub txtStreetLine1_Change()
    If Me.ActiveControl Is txtStreetLine1 Then SeparateFieldsToAddress
End Sub
Private Sub txtStreetLine2_Change()
    If Me.ActiveControl Is txtStreetLine2 Then SeparateFieldsToAddress
End Sub
Private Sub txtCity_Change()
    If Me.ActiveControl Is txtCity Then SeparateFieldsToAddress
End Sub
Private Sub txtState_Change()
    If Me.ActiveControl Is txtState Then SeparateFieldsToAddress
End Sub
Private Sub txtZip_Change()
    If Me.ActiveControl Is txtZip Then SeparateFieldsToAddress
End Sub
Private Sub txtZip4_Change()
    If Me.ActiveControl Is txtZip4 Then SeparateFieldsToAddress
End Sub

Private Sub SeparateFieldsToAddress()
    ctlPostalAddress.StreetLine1 = txtStreetLine1.Text
    ctlPostalAddress.StreetLine2 = txtStreetLine2.Text
    ctlPostalAddress.City = txtCity.Text
    ctlPostalAddress.State = txtState.Text
    ctlPostalAddress.Zip = txtZip.Text
    ctlPostalAddress.Zip4 = txtZip4.Text
End Sub

Private Sub AddressToSeparateFields()
    txtStreetLine1.Text = ctlPostalAddress.StreetLine1
    txtStreetLine2.Text = ctlPostalAddress.StreetLine2
    txtCity.Text = ctlPostalAddress.City
    txtState.Text = ctlPostalAddress.State
    txtZip.Text = ctlPostalAddress.Zip
    txtZip4.Text = ctlPostalAddress.Zip4
End Sub
