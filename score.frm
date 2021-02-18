VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   6210
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form4"
   Picture         =   "score.frx":0000
   ScaleHeight     =   414
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   20
      Top             =   5760
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Index           =   9
      Left            =   2880
      TabIndex        =   19
      Text            =   "Text2"
      Top             =   5160
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Index           =   8
      Left            =   2880
      TabIndex        =   18
      Text            =   "Text2"
      Top             =   4680
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Index           =   7
      Left            =   2880
      TabIndex        =   17
      Text            =   "Text2"
      Top             =   4200
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Index           =   6
      Left            =   2880
      TabIndex        =   16
      Text            =   "Text2"
      Top             =   3720
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Index           =   5
      Left            =   2880
      TabIndex        =   15
      Text            =   "Text2"
      Top             =   3240
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Index           =   4
      Left            =   2880
      TabIndex        =   14
      Text            =   "Text2"
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Index           =   3
      Left            =   2880
      TabIndex        =   13
      Text            =   "Text2"
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Index           =   2
      Left            =   2880
      TabIndex        =   12
      Text            =   "Text2"
      Top             =   1800
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Index           =   1
      Left            =   2880
      TabIndex        =   11
      Text            =   "Text2"
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Index           =   0
      Left            =   2880
      TabIndex        =   10
      Text            =   "Text2"
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   9
      Left            =   240
      TabIndex        =   9
      Top             =   5160
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   8
      Left            =   240
      TabIndex        =   8
      Top             =   4680
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   7
      Left            =   240
      TabIndex        =   7
      Top             =   4200
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   6
      Left            =   240
      TabIndex        =   6
      Top             =   3720
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   5
      Left            =   240
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   3240
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   4
      Left            =   240
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   2760
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Index           =   3
      Left            =   240
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   2280
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   2
      Left            =   240
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1800
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1320
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   840
      Width           =   2415
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Me.Hide
End Sub

Private Sub Form_Load()
Set rs = New Recordset
Set con = New Connection
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\updown.mdb;Persist Security Info=False"
rs.Open "select * from updown order by moves desc", con, adOpenKeyset, adLockOptimistic
For i = 0 To 9
Text1(i).Text = rs!Name
Text2(i).Text = rs!moves
rs.MoveNext
Next
End Sub
