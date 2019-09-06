VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Form3"
   ClientHeight    =   8415
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13650
   LinkTopic       =   "Form3"
   ScaleHeight     =   8415
   ScaleWidth      =   13650
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Show another"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9360
      TabIndex        =   13
      Top             =   5520
      Width           =   3615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9360
      TabIndex        =   12
      Top             =   6480
      Width           =   3615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CLICK HERE TO COMPARE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   9360
      TabIndex        =   11
      Top             =   4200
      Width           =   3615
   End
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   10680
      TabIndex        =   4
      Text            =   "Text5"
      Top             =   3480
      Width           =   2655
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   10680
      TabIndex        =   3
      Text            =   "Text4"
      Top             =   2760
      Width           =   2655
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   10680
      TabIndex        =   2
      Text            =   "Text3"
      Top             =   2040
      Width           =   2655
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   10680
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   1320
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   10680
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   600
      Width           =   2655
   End
   Begin VB.Label Label6 
      Caption         =   "MTR"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8880
      TabIndex        =   10
      Top             =   3480
      Width           =   1695
   End
   Begin VB.Label Label5 
      Caption         =   "Mileage"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8880
      TabIndex        =   9
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   "Fuel"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8880
      TabIndex        =   8
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Engine"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8880
      TabIndex        =   7
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Price"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8880
      TabIndex        =   6
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   8655
   End
   Begin VB.Image Image1 
      Height          =   5775
      Left            =   120
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   8655
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form4.Show
Me.Hide
End Sub
Private Sub Command2_Click()
End
End Sub
Private Sub Command3_Click()
Unload Me
Form2.Show
End Sub
Private Sub Form_Activate()
Text1.Locked = True
Text2.Locked = True
Text3.Locked = True
Text4.Locked = True
Text5.Locked = True
End Sub
