VERSION 5.00
Begin VB.Form Form8 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Form8"
   ClientHeight    =   7800
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14070
   LinkTopic       =   "Form8"
   ScaleHeight     =   7800
   ScaleWidth      =   14070
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text6 
      Height          =   495
      Left            =   10560
      Locked          =   -1  'True
      TabIndex        =   14
      Text            =   "Text6"
      Top             =   240
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   10560
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   960
      Width           =   2655
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   10560
      TabIndex        =   6
      Text            =   "Text2"
      Top             =   1680
      Width           =   2655
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   10560
      TabIndex        =   5
      Text            =   "Text3"
      Top             =   2400
      Width           =   2655
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   10560
      TabIndex        =   4
      Text            =   "Text4"
      Top             =   3120
      Width           =   2655
   End
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   10560
      TabIndex        =   3
      Text            =   "Text5"
      Top             =   3840
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "UPDATE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   9240
      TabIndex        =   2
      Top             =   4560
      Width           =   3615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "LOG OUT"
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
      Left            =   9240
      TabIndex        =   1
      Top             =   6600
      Width           =   3615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Choose another"
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
      Left            =   9240
      TabIndex        =   0
      Top             =   5640
      Width           =   3615
   End
   Begin VB.Label Label7 
      Caption         =   "Model"
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
      Left            =   8760
      TabIndex        =   15
      Top             =   240
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   5775
      Left            =   0
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   8655
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
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   8655
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
      Left            =   8760
      TabIndex        =   12
      Top             =   960
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
      Left            =   8760
      TabIndex        =   11
      Top             =   1680
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
      Left            =   8760
      TabIndex        =   10
      Top             =   2400
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
      Left            =   8760
      TabIndex        =   9
      Top             =   3120
      Width           =   1695
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
      Left            =   8760
      TabIndex        =   8
      Top             =   3840
      Width           =   1695
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CN As ADODB.Connection
Dim RS As ADODB.Recordset
Private Sub Command1_Click()
Dim qry
qry = "Update table1 set Price='" & Text1.Text & "', Engine =  '" & Text2.Text & "', Fuel='" & Text3.Text & "', Mileage='" & Text4.Text & "', MTR='" & Text5.Text & "' where Model ='" & Text6.Text & "';"
CN.Execute qry
MsgBox "SUCCESSFULLY UPDATED"
End Sub
Private Sub Command2_Click()
MsgBox "Logged out successfully", vbInformation, "Sign Out"
End
End Sub
Private Sub Command3_Click()
Unload Me
Unload Form7
Form7.Show
End Sub
Private Sub Form_Activate()
Set CN = New ADODB.Connection
CN.ConnectionString = "Provider = Microsoft.Jet.OLEDB.4.0; Data Source = vehicle.mdb;"
CN.Open
If CN.State = 1 Then
'a = MsgBox("A successful connection is established", vbInformation, "Message")
End If
End Sub
