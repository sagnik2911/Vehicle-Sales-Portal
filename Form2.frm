VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form2 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Form2"
   ClientHeight    =   4290
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7620
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   4290
   ScaleWidth      =   7620
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   120
      Top             =   3240
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\Studies\VB project\vehicle.mdb;Mode=Share Deny None;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\Studies\VB project\vehicle.mdb;Mode=Share Deny None;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2040
      TabIndex        =   3
      Top             =   2640
      Width           =   3375
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   3960
      TabIndex        =   2
      Text            =   "Combo2"
      Top             =   1680
      Width           =   3615
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   120
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   1680
      Width           =   3255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Niagara Solid"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   7215
   End
End
Attribute VB_Name = "Form2"
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
Private Sub Command1_Click()
Set RS = New ADODB.Recordset
RS.Open "select Price, Engine, Fuel, Mileage, MTR,link from table1 where Model = '" & Combo2.Text & "'", CN, adOpenDynamic
Form3.Show
Form3.Label1.Caption = Combo1.Text & "-" & Combo2.Text
Form3.Text1.Text = RS(0).Value
Form3.Text2.Text = RS(1).Value
Form3.Text3.Text = RS(2).Value
Form3.Text4.Text = RS(3).Value
Form3.Text5.Text = RS(4).Value
Form5.Label1.Caption = Combo1.Text & "-" & Combo2.Text
Form5.Text6.Text = RS(0).Value
Form5.Text7.Text = RS(1).Value
Form5.Text8.Text = RS(2).Value
Form5.Text9.Text = RS(3).Value
Form5.Text10.Text = RS(4).Value
Form3.Image1.Picture = LoadPicture(RS(5).Value)
Unload Me
End Sub
