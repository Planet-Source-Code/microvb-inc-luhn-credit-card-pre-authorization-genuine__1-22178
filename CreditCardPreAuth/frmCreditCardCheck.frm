VERSION 5.00
Begin VB.Form frmCreditCardCheck 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "LUHN Credit Pre-Authorization"
   ClientHeight    =   2700
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5070
   Icon            =   "frmCreditCardCheck.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   5070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   570
      Index           =   1
      Left            =   810
      TabIndex        =   4
      Top             =   1920
      Visible         =   0   'False
      Width           =   3330
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "VALID !!!"
         ForeColor       =   &H00004000&
         Height          =   195
         Index           =   3
         Left            =   150
         TabIndex        =   5
         Top             =   225
         Width           =   2970
      End
      Begin VB.Shape shp 
         FillColor       =   &H00C0FFC0&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   2
         Left            =   105
         Top             =   195
         Width           =   3075
      End
      Begin VB.Shape shp 
         BorderColor     =   &H00E0E0E0&
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   3
         Left            =   120
         Top             =   240
         Width           =   3105
      End
   End
   Begin VB.TextBox txtCreditCardNumber 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H00404000&
      Height          =   285
      Left            =   1695
      MaxLength       =   20
      TabIndex        =   2
      Top             =   1320
      Width           =   2505
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   570
      Index           =   0
      Left            =   825
      TabIndex        =   0
      Top             =   1875
      Visible         =   0   'False
      Width           =   3330
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "INVALID !!!"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   1
         Top             =   225
         Width           =   2970
      End
      Begin VB.Shape shp 
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   0
         Left            =   105
         Top             =   195
         Width           =   3075
      End
      Begin VB.Shape shp 
         BorderColor     =   &H00E0E0E0&
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   1
         Left            =   135
         Top             =   240
         Width           =   3105
      End
   End
   Begin VB.Label Label1 
      Caption         =   $"frmCreditCardCheck.frx":1042
      Height          =   990
      Index           =   7
      Left            =   225
      TabIndex        =   10
      Top             =   135
      Width           =   4635
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   195
      Index           =   6
      Left            =   1590
      TabIndex        =   9
      Top             =   2505
      Width           =   3405
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Checksum Value :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   195
      Index           =   5
      Left            =   15
      TabIndex        =   8
      Top             =   2505
      Width           =   1545
   End
   Begin VB.Shape shp 
      BackColor       =   &H00808000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      FillColor       =   &H00C0C000&
      FillStyle       =   7  'Diagonal Cross
      Height          =   225
      Index           =   4
      Left            =   0
      Top             =   2490
      Width           =   5070
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   4
      Left            =   1740
      TabIndex        =   7
      Top             =   1665
      Width           =   2550
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Credit Card :"
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   2
      Left            =   630
      TabIndex        =   6
      Top             =   1695
      Width           =   1035
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Credit Card # :"
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   1
      Left            =   630
      TabIndex        =   3
      Top             =   1350
      Width           =   1035
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Please type a number above."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   255
      Index           =   8
      Left            =   975
      TabIndex        =   11
      Top             =   2055
      Width           =   3075
   End
End
Attribute VB_Name = "frmCreditCardCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub txtCreditCardNumber_Change()
  If txtCreditCardNumber <> "" Then
    Dim ccReturn As CreditCardStats
    ccReturn = ValidateCreditCardNumber(txtCreditCardNumber)
    'Check the credit card company. If the credit card company is ""
    'then the card is deemed invalid. With some minor changes to the
    'ValidateCreditCardNumber function in modCCPreAuth, you can change
    'this so that it will still return the CreditCardNumber, and use
    'the ccReturn.IsValidNumber boolean type data to return TRUE or FALSE
    'depending on if the card is valid or not respectively.
    Select Case ccReturn.CreditCardCo
      Case 0 ' INVALID
          Frame1(0).Visible = True
          Frame1(1).Visible = False
          Label1(4).Caption = ""
          Label1(6).Caption = " N / A"
      Case 1 ' Mastercard
          Frame1(0).Visible = False
          Frame1(1).Visible = True
          Label1(4).Caption = "MASTERCARD"
          'display the checksum validity data
          Label1(6).Caption = ccReturn.CheckSum
      Case 2 ' VISA
          Frame1(0).Visible = False
          Frame1(1).Visible = True
          Label1(4).Caption = "VISA"
          'display the checksum validity data
          Label1(6).Caption = ccReturn.CheckSum
      Case 3 ' American Express
          Frame1(0).Visible = False
          Frame1(1).Visible = True
          Label1(4).Caption = "AMERICAN EXPRESS"
          'display the checksum validity data
          Label1(6).Caption = ccReturn.CheckSum
      Case 4 ' Diners Club / Carte Blanche
          Frame1(0).Visible = False
          Frame1(1).Visible = True
          Label1(4).Caption = "DINERS CLUB / CARTE BLANCHE"
          'display the checksum validity data
          Label1(6).Caption = ccReturn.CheckSum
      Case 5 ' Discover
          Frame1(0).Visible = False
          Frame1(1).Visible = True
          Label1(4).Caption = "DISCOVER"
          'display the checksum validity data
          Label1(6).Caption = ccReturn.CheckSum
      Case 6 ' enRoute
          Frame1(0).Visible = False
          Frame1(1).Visible = True
          Label1(4).Caption = "enROUTE"
          'display the checksum validity data
          Label1(6).Caption = ccReturn.CheckSum
      Case 7 ' JCB
          Frame1(0).Visible = False
          Frame1(1).Visible = True
          Label1(4).Caption = "JCB"
          'display the checksum validity data
          Label1(6).Caption = ccReturn.CheckSum
    End Select
  Else
    'Just some simple code to make the form look
    'better when there is nothing in the text box.
    Frame1(0).Visible = False
    Frame1(1).Visible = False
    Label1(4).Caption = ""
    Label1(6).Caption = ""
  End If
End Sub
