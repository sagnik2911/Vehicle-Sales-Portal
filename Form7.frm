VERSION 5.00
Begin VB.Form Form7 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Form7"
   ClientHeight    =   4005
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5655
   FillColor       =   &H8000000B&
   LinkTopic       =   "Form7"
   ScaleHeight     =   4005
   ScaleWidth      =   5655
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Model"
      Height          =   1935
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   3375
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   360
         TabIndex        =   3
         Top             =   960
         Width           =   2535
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   360
         TabIndex        =   2
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "CHOOSE THE MODEL"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4935
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CN As ADODB.Connection
Dim RS As ADODB.Recordset
Private Sub Combo1_Click()
Dim a As String
If Combo1.Text = "Hyundai" Then
Open "C:\Users\user\Desktop\VB project\hyundai.txt" For Input As #1
While EOF(1) = False
    Line Input #1, a
    Combo2.AddItem a
Wend
Close #1
ElseIf Combo1.Text = "Maruti" Then
Open "C:\Users\user\Desktop\VB project\maruti.txt" For Input As #1
While EOF(1) = False
    Line Input #1, a
    Combo2.AddItem a
Wend
Close #1
ElseIf Combo1.Text = "Chevrolet" Then
Open "C:\Users\user\Desktop\VB project\chevrolet.txt" For Input As #1
While EOF(1) = False
    Line Input #1, a
    Combo2.AddItem a
Wend
Close #1
End If
Combo1.Enabled = False
End Sub
Private Sub Combo2_Click()
Set RS = New ADODB.Recordset
RS.Open "select Price, Engine, Fuel, Mileage, MTR,link from table1 where Model = '" & Combo2.Text & "'", CN, adOpenDynamic
Form8.Show
Form8.Label1.Caption = Combo1.Text & "-" & Combo2.Text
Form8.Text1.Text = RS(0).Value
Form8.Text2.Text = RS(1).Value
Form8.Text3.Text = RS(2).Value
Form8.Text4.Text = RS(3).Value
Form8.Text5.Text = RS(4).Value
Form8.Text6.Text = Combo2.Text
Form8.Image1.Picture = LoadPicture(RS(5).Value)
Me.Hide
End Sub
Private Sub Form_Load()
Label1.Caption = "Choose your Vehicle"
Combo1.Text = "Makers"
Combo1.AddItem "Hyundai"
Combo1.AddItem "Maruti"
Combo1.AddItem "Chevrolet"
Combo2.Text = "Model"
End Sub
Private Sub Form_Activate()
Set CN = New ADODB.Connection
CN.ConnectionString = "Provider = Microsoft.Jet.OLEDB.4.0; Data Source = vehicle.mdb;"
CN.Open
If CN.State = 1 Then
'a = MsgBox("A successful connection is established", vbInformation, "Message")
End If
End Sub
