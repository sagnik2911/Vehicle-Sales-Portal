VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Form4"
   ClientHeight    =   3930
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7875
   LinkTopic       =   "Form4"
   ScaleHeight     =   3930
   ScaleWidth      =   7875
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Okay"
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
      Left            =   120
      TabIndex        =   3
      Top             =   2280
      Width           =   2415
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Text            =   "Combo2"
      Top             =   1560
      Width           =   2295
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   840
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   3375
      Left            =   2760
      Stretch         =   -1  'True
      Top             =   120
      Width           =   4935
   End
   Begin VB.Label Label1 
      Caption         =   "Compare with:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "Form4"
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
RS.Open "select link from table1 where Model = '" & Combo2.Text & "'", CN, adOpenDynamic
Image1.Picture = LoadPicture(RS(0).Value)
End Sub

Private Sub Command1_Click()
Set RS = New ADODB.Recordset
RS.Open "select Price, Engine, Fuel, Mileage, MTR from table1 where Model = '" & Combo2.Text & "'", CN, adOpenDynamic
Form5.Label2.Caption = Combo1.Text & "-" & Combo2.Text
Form5.Text1.Text = RS(0).Value
Form5.Text2.Text = RS(1).Value
Form5.Text3.Text = RS(2).Value
Form5.Text4.Text = RS(3).Value
Form5.Text5.Text = RS(4).Value
Unload Form4
Me.Hide
Form5.Show
End Sub
Private Sub Form_Load()
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
'If CN.State = 1 Then
'a = MsgBox("A successful connection is established", vbInformation, "Message")
'End If
End Sub

