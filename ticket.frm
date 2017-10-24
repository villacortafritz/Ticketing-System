VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "To:"
   ClientHeight    =   5340
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3465
   LinkTopic       =   "Form1"
   ScaleHeight     =   5340
   ScaleWidth      =   3465
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox dropTo 
      Height          =   315
      ItemData        =   "ticket.frx":0000
      Left            =   1080
      List            =   "ticket.frx":0025
      TabIndex        =   17
      Top             =   2040
      Width           =   2175
   End
   Begin VB.TextBox labelFare 
      Height          =   405
      Left            =   1080
      TabIndex        =   16
      Text            =   "PHP"
      Top             =   4680
      Width           =   2055
   End
   Begin VB.TextBox textInfant 
      Height          =   285
      Left            =   2280
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   3000
      Width           =   735
   End
   Begin VB.TextBox textPassenger 
      Height          =   285
      Left            =   2280
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   2520
      Width           =   735
   End
   Begin VB.ComboBox dropKilos 
      Height          =   315
      Left            =   2520
      Sorted          =   -1  'True
      TabIndex        =   11
      Top             =   3480
      Width           =   735
   End
   Begin VB.CheckBox checkSports 
      Caption         =   "Sports Utility"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   4200
      Width           =   1575
   End
   Begin VB.CheckBox checkMeals 
      Caption         =   "Meals"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   3840
      Width           =   1335
   End
   Begin VB.CheckBox checkBaggage 
      Caption         =   "Baggage"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   3480
      Width           =   1095
   End
   Begin VB.ComboBox dropFrom 
      Height          =   315
      ItemData        =   "ticket.frx":00C7
      Left            =   1080
      List            =   "ticket.frx":00EC
      TabIndex        =   5
      Top             =   1560
      Width           =   2175
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Vice Versa"
      Height          =   195
      Left            =   1080
      TabIndex        =   2
      Top             =   600
      Width           =   2175
   End
   Begin VB.OptionButton Option1 
      Caption         =   "One Way"
      Height          =   195
      Left            =   1080
      TabIndex        =   1
      Top             =   960
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select"
      Height          =   1095
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3015
   End
   Begin VB.Line Line4 
      X1              =   240
      X2              =   3240
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Line Line3 
      X1              =   240
      X2              =   3240
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line2 
      X1              =   240
      X2              =   3240
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   3240
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Label Label6 
      Caption         =   "FARE:"
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
      Left            =   240
      TabIndex        =   13
      Top             =   4680
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "No. of Kilos: "
      Height          =   255
      Left            =   1440
      TabIndex        =   12
      Top             =   3480
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "No. of Infant Passengers:"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "No. of Passengers:"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "To:"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "From:"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1560
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fee As Double




Private Sub labelFare_Change()
    If dropFrom = "Privet Drive" Then
        If dropTo = "Grimmauld Place" Then fee = 50
        ElseIf dropTo = "King's Cross" Then fee = 100
        ElseIf dropTo = "Godric's Hollow" Then fee = 150
        ElseIf dropTo = "Ministry of Magic" Then fee = 150
        ElseIf dropTo = "Diagon Alley" Then fee = 250
        ElseIf dropTo = "Knockturn Alley" Then fee = 300
        ElseIf dropTo = "Burrow" Then fee = 400
        ElseIf dropTo = "Hogwarts" Then fee = 500
        ElseIf dropTo = "Shrieking Shack" Then fee = 600
        ElseIf dropTo = "Hogsmeade" Then fee = 550
        
    ElseIf dropFrom = "King's Cross" Then
        If dropTo = "Grimmauld Place" Then fee = 50
        ElseIf dropTo = "Privet Drive" Then fee = 100
        ElseIf dropTo = "Godric's Hollow" Then fee = 50
        ElseIf dropTo = "Ministry of Magic" Then fee = 150
        ElseIf dropTo = "Diagon Alley" Then fee = 250
        ElseIf dropTo = "Knockturn Alley" Then fee = 300
        ElseIf dropTo = "Burrow" Then fee = 400
        ElseIf dropTo = "Hogwarts" Then fee = 500
        ElseIf dropTo = "Shrieking Shack" Then fee = 600
        ElseIf dropTo = "Hogsmeade" Then fee = 550
    End If
    
    
End Sub


Private Sub Option1_Click()
    labelFare.Text = fare * 2
End Sub

