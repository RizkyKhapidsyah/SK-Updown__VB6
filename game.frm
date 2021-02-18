VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "SWFLASH.OCX"
Begin VB.Form Form1 
   BackColor       =   &H000040C0&
   ClientHeight    =   7545
   ClientLeft      =   -2145
   ClientTop       =   -1545
   ClientWidth     =   11400
   LinkTopic       =   "Form1"
   ScaleHeight     =   8778.359
   ScaleMode       =   0  'User
   ScaleWidth      =   11532.62
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9360
      TabIndex        =   107
      Top             =   6480
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "High Score"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8640
      TabIndex        =   106
      Top             =   5880
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8640
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   3000
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Move"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10080
      TabIndex        =   2
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Shuffle"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8640
      TabIndex        =   1
      Top             =   2520
      Width           =   1215
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash ShockwaveFlash1 
      Height          =   2175
      Left            =   9120
      TabIndex        =   0
      Top             =   240
      Width           =   2175
      _cx             =   4198140
      _cy             =   4198140
      Movie           =   "c:\program files\microsoft visual studio\vb98\ladder&snake\dice2.swf"
      Src             =   "c:\program files\microsoft visual studio\vb98\ladder&snake\dice2.swf"
      WMode           =   "Window"
      Play            =   0   'False
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   -1  'True
      Base            =   ""
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
      Stacking        =   "below"
   End
   Begin VB.Image Image3 
      Height          =   405
      Left            =   8640
      Picture         =   "game.frx":0000
      Top             =   6600
      Width           =   435
   End
   Begin VB.Label Label3 
      BackColor       =   &H000040C0&
      Caption         =   "<--Go DownWards"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9720
      TabIndex        =   105
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Line Line59 
      BorderWidth     =   3
      X1              =   9590.284
      X2              =   9711.681
      Y1              =   6143.106
      Y2              =   6282.723
   End
   Begin VB.Line Line58 
      BorderWidth     =   3
      X1              =   9711.681
      X2              =   9833.076
      Y1              =   6282.723
      Y2              =   6143.106
   End
   Begin VB.Line Line57 
      BorderWidth     =   3
      X1              =   9711.681
      X2              =   9711.681
      Y1              =   5026.178
      Y2              =   6282.723
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Index           =   59
      Left            =   7680
      Top             =   3120
      Width           =   855
   End
   Begin VB.Line Line56 
      BorderWidth     =   3
      X1              =   1456.752
      X2              =   1335.356
      Y1              =   4886.562
      Y2              =   4746.946
   End
   Begin VB.Line Line55 
      BorderWidth     =   3
      X1              =   1456.752
      X2              =   1578.148
      Y1              =   4886.562
      Y2              =   4746.946
   End
   Begin VB.Line Line17 
      BorderWidth     =   3
      X1              =   2063.732
      X2              =   1456.752
      Y1              =   977.312
      Y2              =   4886.562
   End
   Begin VB.Line Line12 
      BorderWidth     =   3
      X1              =   4127.464
      X2              =   4127.464
      Y1              =   6282.723
      Y2              =   6422.339
   End
   Begin VB.Line Line11 
      BorderWidth     =   3
      X1              =   4127.464
      X2              =   4248.86
      Y1              =   6422.339
      Y2              =   6422.339
   End
   Begin VB.Line Line10 
      BorderWidth     =   3
      X1              =   4127.464
      X2              =   5341.424
      Y1              =   6422.339
      Y2              =   4188.482
   End
   Begin VB.Line Line19 
      BorderWidth     =   3
      Index           =   12
      X1              =   5341.424
      X2              =   5341.424
      Y1              =   5863.875
      Y2              =   7539.267
   End
   Begin VB.Line Line20 
      BorderWidth     =   3
      Index           =   12
      X1              =   5584.216
      X2              =   5584.216
      Y1              =   5863.875
      Y2              =   7539.267
   End
   Begin VB.Line Line21 
      BorderWidth     =   3
      Index           =   12
      X1              =   5341.424
      X2              =   5584.216
      Y1              =   6282.723
      Y2              =   6282.723
   End
   Begin VB.Line Line22 
      BorderWidth     =   3
      Index           =   6
      X1              =   5341.424
      X2              =   5584.216
      Y1              =   6561.955
      Y2              =   6561.955
   End
   Begin VB.Line Line23 
      BorderWidth     =   3
      Index           =   12
      X1              =   5341.424
      X2              =   5584.216
      Y1              =   7120.419
      Y2              =   7120.419
   End
   Begin VB.Line Line24 
      BorderWidth     =   3
      Index           =   6
      X1              =   5341.424
      X2              =   5584.216
      Y1              =   6003.49
      Y2              =   6003.49
   End
   Begin VB.Line Line25 
      BorderWidth     =   3
      Index           =   12
      X1              =   5341.424
      X2              =   5584.216
      Y1              =   6841.187
      Y2              =   6841.187
   End
   Begin VB.Line Line26 
      BorderWidth     =   3
      Index           =   6
      X1              =   5341.424
      X2              =   5584.216
      Y1              =   7399.651
      Y2              =   7399.651
   End
   Begin VB.Line Line54 
      BorderWidth     =   3
      X1              =   6919.572
      X2              =   7040.968
      Y1              =   7260.035
      Y2              =   7399.651
   End
   Begin VB.Line Line53 
      BorderWidth     =   3
      X1              =   7040.968
      X2              =   7162.364
      Y1              =   7399.651
      Y2              =   7260.035
   End
   Begin VB.Line Line52 
      BorderWidth     =   3
      X1              =   7162.364
      X2              =   7040.968
      Y1              =   5165.794
      Y2              =   7399.651
   End
   Begin VB.Line Line51 
      BorderWidth     =   3
      X1              =   1578.148
      X2              =   1456.752
      Y1              =   8097.731
      Y2              =   7958.115
   End
   Begin VB.Line Line50 
      BorderWidth     =   3
      X1              =   1578.148
      X2              =   1699.544
      Y1              =   8097.731
      Y2              =   7958.115
   End
   Begin VB.Line Line49 
      BorderWidth     =   3
      X1              =   2063.732
      X2              =   1578.148
      Y1              =   6003.49
      Y2              =   8097.731
   End
   Begin VB.Line Line48 
      BorderWidth     =   3
      X1              =   3641.88
      X2              =   3399.088
      Y1              =   2931.937
      Y2              =   2931.937
   End
   Begin VB.Line Line47 
      BorderWidth     =   3
      X1              =   3641.88
      X2              =   3641.88
      Y1              =   2931.937
      Y2              =   2652.705
   End
   Begin VB.Line Line46 
      BorderWidth     =   3
      X1              =   2549.316
      X2              =   3641.88
      Y1              =   1815.009
      Y2              =   2931.937
   End
   Begin VB.Line Line45 
      BorderWidth     =   3
      X1              =   4006.068
      X2              =   4248.86
      Y1              =   7958.115
      Y2              =   8097.731
   End
   Begin VB.Line Line44 
      BorderWidth     =   3
      X1              =   4248.86
      X2              =   4370.256
      Y1              =   8097.731
      Y2              =   7818.499
   End
   Begin VB.Line Line43 
      BorderWidth     =   3
      X1              =   3884.672
      X2              =   4248.86
      Y1              =   4188.482
      Y2              =   8097.731
   End
   Begin VB.Line Line42 
      BorderWidth     =   3
      X1              =   4977.236
      X2              =   4977.236
      Y1              =   3769.634
      Y2              =   3490.401
   End
   Begin VB.Line Line41 
      BorderWidth     =   3
      X1              =   4977.236
      X2              =   5220.028
      Y1              =   3769.634
      Y2              =   3769.634
   End
   Begin VB.Line Line40 
      BorderWidth     =   3
      X1              =   7162.364
      X2              =   4977.236
      Y1              =   837.696
      Y2              =   3769.634
   End
   Begin VB.Line Line39 
      BorderWidth     =   3
      X1              =   485.584
      X2              =   606.98
      Y1              =   7958.115
      Y2              =   8097.731
   End
   Begin VB.Line Line38 
      BorderWidth     =   3
      X1              =   728.376
      X2              =   606.98
      Y1              =   7958.115
      Y2              =   8097.731
   End
   Begin VB.Line Line37 
      BorderWidth     =   3
      X1              =   242.792
      X2              =   606.98
      Y1              =   5584.642
      Y2              =   8097.731
   End
   Begin VB.Line Line36 
      BorderWidth     =   3
      X1              =   6555.384
      X2              =   6676.78
      Y1              =   3769.634
      Y2              =   4048.866
   End
   Begin VB.Line Line35 
      BorderWidth     =   3
      X1              =   6676.78
      X2              =   6798.176
      Y1              =   4048.866
      Y2              =   3769.634
   End
   Begin VB.Line Line34 
      BorderWidth     =   3
      X1              =   6676.78
      X2              =   6676.78
      Y1              =   2652.705
      Y2              =   4048.866
   End
   Begin VB.Label Label2 
      BackColor       =   &H000040C0&
      Caption         =   "<--Go UpWards"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9000
      TabIndex        =   104
      Top             =   3840
      Width           =   2175
   End
   Begin VB.Line Line26 
      BorderWidth     =   3
      Index           =   12
      X1              =   7162.364
      X2              =   7405.156
      Y1              =   3071.553
      Y2              =   3071.553
   End
   Begin VB.Line Line25 
      BorderWidth     =   3
      Index           =   13
      X1              =   7162.364
      X2              =   7405.156
      Y1              =   2513.089
      Y2              =   2513.089
   End
   Begin VB.Line Line24 
      BorderWidth     =   3
      Index           =   13
      X1              =   7162.364
      X2              =   7405.156
      Y1              =   1675.393
      Y2              =   1675.393
   End
   Begin VB.Line Line23 
      BorderWidth     =   3
      Index           =   13
      X1              =   7162.364
      X2              =   7405.156
      Y1              =   2792.321
      Y2              =   2792.321
   End
   Begin VB.Line Line22 
      BorderWidth     =   3
      Index           =   13
      X1              =   7162.364
      X2              =   7405.156
      Y1              =   2233.857
      Y2              =   2233.857
   End
   Begin VB.Line Line21 
      BorderWidth     =   3
      Index           =   14
      X1              =   7162.364
      X2              =   7405.156
      Y1              =   1954.625
      Y2              =   1954.625
   End
   Begin VB.Line Line20 
      BorderWidth     =   3
      Index           =   14
      X1              =   7405.156
      X2              =   7405.156
      Y1              =   1535.777
      Y2              =   3211.169
   End
   Begin VB.Line Line19 
      BorderWidth     =   3
      Index           =   14
      X1              =   7162.364
      X2              =   7162.364
      Y1              =   1535.777
      Y2              =   3211.169
   End
   Begin VB.Line Line24 
      BorderWidth     =   3
      Index           =   12
      X1              =   3884.672
      X2              =   4127.464
      Y1              =   837.696
      Y2              =   837.696
   End
   Begin VB.Line Line22 
      BorderWidth     =   3
      Index           =   12
      X1              =   3641.88
      X2              =   3884.672
      Y1              =   1396.161
      Y2              =   1396.161
   End
   Begin VB.Line Line21 
      BorderWidth     =   3
      Index           =   13
      X1              =   3763.276
      X2              =   4006.068
      Y1              =   1116.928
      Y2              =   1116.928
   End
   Begin VB.Line Line20 
      BorderWidth     =   3
      Index           =   13
      X1              =   4127.464
      X2              =   3884.672
      Y1              =   698.08
      Y2              =   1535.777
   End
   Begin VB.Line Line19 
      BorderWidth     =   3
      Index           =   13
      X1              =   3884.672
      X2              =   3641.88
      Y1              =   698.08
      Y2              =   1535.777
   End
   Begin VB.Line Line26 
      BorderWidth     =   3
      Index           =   11
      X1              =   7890.74
      X2              =   8133.532
      Y1              =   4746.946
      Y2              =   4746.946
   End
   Begin VB.Line Line25 
      BorderWidth     =   3
      Index           =   11
      X1              =   7890.74
      X2              =   8133.532
      Y1              =   4188.482
      Y2              =   4188.482
   End
   Begin VB.Line Line24 
      BorderWidth     =   3
      Index           =   11
      X1              =   7890.74
      X2              =   8133.532
      Y1              =   3350.785
      Y2              =   3350.785
   End
   Begin VB.Line Line23 
      BorderWidth     =   3
      Index           =   11
      X1              =   7890.74
      X2              =   8133.532
      Y1              =   4467.714
      Y2              =   4467.714
   End
   Begin VB.Line Line22 
      BorderWidth     =   3
      Index           =   11
      X1              =   7890.74
      X2              =   8133.532
      Y1              =   3909.25
      Y2              =   3909.25
   End
   Begin VB.Line Line21 
      BorderWidth     =   3
      Index           =   11
      X1              =   7890.74
      X2              =   8133.532
      Y1              =   3630.018
      Y2              =   3630.018
   End
   Begin VB.Line Line20 
      BorderWidth     =   3
      Index           =   11
      X1              =   8133.532
      X2              =   8133.532
      Y1              =   3211.169
      Y2              =   4886.562
   End
   Begin VB.Line Line19 
      BorderWidth     =   3
      Index           =   11
      X1              =   7890.74
      X2              =   7890.74
      Y1              =   3211.169
      Y2              =   4886.562
   End
   Begin VB.Line Line33 
      BorderWidth     =   3
      X1              =   7769.344
      X2              =   8012.136
      Y1              =   6422.339
      Y2              =   6422.339
   End
   Begin VB.Line Line32 
      BorderWidth     =   3
      X1              =   7526.552
      X2              =   7769.344
      Y1              =   6282.723
      Y2              =   6282.723
   End
   Begin VB.Line Line15 
      BorderWidth     =   3
      X1              =   7405.156
      X2              =   7647.948
      Y1              =   6143.106
      Y2              =   6143.106
   End
   Begin VB.Line Line26 
      BorderWidth     =   3
      Index           =   10
      X1              =   7283.76
      X2              =   7526.552
      Y1              =   6003.49
      Y2              =   6003.49
   End
   Begin VB.Line Line25 
      BorderWidth     =   3
      Index           =   10
      X1              =   6919.572
      X2              =   7162.364
      Y1              =   5724.258
      Y2              =   5724.258
   End
   Begin VB.Line Line24 
      BorderWidth     =   3
      Index           =   10
      X1              =   6433.988
      X2              =   6676.78
      Y1              =   5305.41
      Y2              =   5305.41
   End
   Begin VB.Line Line23 
      BorderWidth     =   3
      Index           =   10
      X1              =   7162.364
      X2              =   7405.156
      Y1              =   5863.875
      Y2              =   5863.875
   End
   Begin VB.Line Line22 
      BorderWidth     =   3
      Index           =   10
      X1              =   6798.176
      X2              =   7040.968
      Y1              =   5584.642
      Y2              =   5584.642
   End
   Begin VB.Line Line21 
      BorderWidth     =   3
      Index           =   10
      X1              =   6676.78
      X2              =   6919.572
      Y1              =   5445.026
      Y2              =   5445.026
   End
   Begin VB.Line Line20 
      BorderWidth     =   3
      Index           =   10
      X1              =   6433.988
      X2              =   8133.532
      Y1              =   5026.178
      Y2              =   6561.955
   End
   Begin VB.Line Line19 
      BorderWidth     =   3
      Index           =   10
      X1              =   6312.592
      X2              =   7890.74
      Y1              =   5165.794
      Y2              =   6561.955
   End
   Begin VB.Line Line26 
      BorderWidth     =   3
      Index           =   9
      X1              =   5341.424
      X2              =   5584.216
      Y1              =   2233.857
      Y2              =   2233.857
   End
   Begin VB.Line Line25 
      BorderWidth     =   3
      Index           =   9
      X1              =   5341.424
      X2              =   5584.216
      Y1              =   1675.393
      Y2              =   1675.393
   End
   Begin VB.Line Line24 
      BorderWidth     =   3
      Index           =   9
      X1              =   5341.424
      X2              =   5584.216
      Y1              =   837.696
      Y2              =   837.696
   End
   Begin VB.Line Line23 
      BorderWidth     =   3
      Index           =   9
      X1              =   5341.424
      X2              =   5584.216
      Y1              =   1954.625
      Y2              =   1954.625
   End
   Begin VB.Line Line22 
      BorderWidth     =   3
      Index           =   9
      X1              =   5341.424
      X2              =   5584.216
      Y1              =   1396.161
      Y2              =   1396.161
   End
   Begin VB.Line Line21 
      BorderWidth     =   3
      Index           =   9
      X1              =   5341.424
      X2              =   5584.216
      Y1              =   1116.928
      Y2              =   1116.928
   End
   Begin VB.Line Line20 
      BorderWidth     =   3
      Index           =   9
      X1              =   5584.216
      X2              =   5584.216
      Y1              =   698.08
      Y2              =   2373.473
   End
   Begin VB.Line Line19 
      BorderWidth     =   3
      Index           =   9
      X1              =   5341.424
      X2              =   5341.424
      Y1              =   698.08
      Y2              =   2373.473
   End
   Begin VB.Line Line14 
      BorderWidth     =   3
      X1              =   2306.524
      X2              =   2549.316
      Y1              =   5305.41
      Y2              =   5305.41
   End
   Begin VB.Line Line13 
      BorderWidth     =   3
      X1              =   2427.92
      X2              =   2670.712
      Y1              =   5026.178
      Y2              =   5026.178
   End
   Begin VB.Line Line26 
      BorderWidth     =   3
      Index           =   8
      X1              =   2549.316
      X2              =   2792.108
      Y1              =   4746.946
      Y2              =   4746.946
   End
   Begin VB.Line Line25 
      BorderWidth     =   3
      Index           =   8
      X1              =   2792.108
      X2              =   3034.9
      Y1              =   4188.482
      Y2              =   4188.482
   End
   Begin VB.Line Line24 
      BorderWidth     =   3
      Index           =   8
      X1              =   3156.296
      X2              =   3399.088
      Y1              =   3350.785
      Y2              =   3350.785
   End
   Begin VB.Line Line23 
      BorderWidth     =   3
      Index           =   8
      X1              =   2670.712
      X2              =   2913.504
      Y1              =   4467.714
      Y2              =   4467.714
   End
   Begin VB.Line Line22 
      BorderWidth     =   3
      Index           =   8
      X1              =   2913.504
      X2              =   3156.296
      Y1              =   3909.25
      Y2              =   3909.25
   End
   Begin VB.Line Line21 
      BorderWidth     =   3
      Index           =   8
      X1              =   3034.9
      X2              =   3277.692
      Y1              =   3630.018
      Y2              =   3630.018
   End
   Begin VB.Line Line20 
      BorderWidth     =   3
      Index           =   8
      X1              =   3520.484
      X2              =   2427.92
      Y1              =   3211.169
      Y2              =   5584.642
   End
   Begin VB.Line Line19 
      BorderWidth     =   3
      Index           =   8
      X1              =   3277.692
      X2              =   2185.128
      Y1              =   3071.553
      Y2              =   5584.642
   End
   Begin VB.Line Line26 
      BorderWidth     =   3
      Index           =   7
      X1              =   1942.336
      X2              =   2185.128
      Y1              =   4607.33
      Y2              =   4607.33
   End
   Begin VB.Line Line25 
      BorderWidth     =   3
      Index           =   7
      X1              =   1942.336
      X2              =   2185.128
      Y1              =   4048.866
      Y2              =   4048.866
   End
   Begin VB.Line Line24 
      BorderWidth     =   3
      Index           =   7
      X1              =   1942.336
      X2              =   2185.128
      Y1              =   3211.169
      Y2              =   3211.169
   End
   Begin VB.Line Line23 
      BorderWidth     =   3
      Index           =   7
      X1              =   1942.336
      X2              =   2185.128
      Y1              =   4328.098
      Y2              =   4328.098
   End
   Begin VB.Line Line22 
      BorderWidth     =   3
      Index           =   7
      X1              =   1942.336
      X2              =   2185.128
      Y1              =   3769.634
      Y2              =   3769.634
   End
   Begin VB.Line Line21 
      BorderWidth     =   3
      Index           =   7
      X1              =   1942.336
      X2              =   2185.128
      Y1              =   3490.401
      Y2              =   3490.401
   End
   Begin VB.Line Line20 
      BorderWidth     =   3
      Index           =   7
      X1              =   2185.128
      X2              =   2185.128
      Y1              =   3071.553
      Y2              =   4746.946
   End
   Begin VB.Line Line19 
      BorderWidth     =   3
      Index           =   7
      X1              =   1942.336
      X2              =   1942.336
      Y1              =   3071.553
      Y2              =   4746.946
   End
   Begin VB.Line Line25 
      BorderWidth     =   3
      Index           =   6
      X1              =   5584.216
      X2              =   5827.008
      Y1              =   3769.634
      Y2              =   3769.634
   End
   Begin VB.Line Line23 
      BorderWidth     =   3
      Index           =   6
      X1              =   5948.404
      X2              =   6191.196
      Y1              =   3490.401
      Y2              =   3490.401
   End
   Begin VB.Line Line21 
      BorderWidth     =   3
      Index           =   6
      X1              =   5827.008
      X2              =   6069.8
      Y1              =   3630.018
      Y2              =   3630.018
   End
   Begin VB.Line Line20 
      BorderWidth     =   3
      Index           =   6
      X1              =   6312.592
      X2              =   5584.216
      Y1              =   3490.401
      Y2              =   3909.25
   End
   Begin VB.Line Line19 
      BorderWidth     =   3
      Index           =   6
      X1              =   6191.196
      X2              =   5462.82
      Y1              =   3350.785
      Y2              =   3769.634
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   66
      Left            =   3000
      TabIndex        =   103
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   54
      Left            =   3840
      TabIndex        =   102
      Top             =   3360
      Width           =   375
   End
   Begin VB.Line Line31 
      BorderStyle     =   2  'Dash
      BorderWidth     =   3
      X1              =   606.98
      X2              =   849.772
      Y1              =   3630.018
      Y2              =   3630.018
   End
   Begin VB.Line Line30 
      BorderStyle     =   2  'Dash
      BorderWidth     =   3
      X1              =   606.98
      X2              =   849.772
      Y1              =   3350.785
      Y2              =   3350.785
   End
   Begin VB.Line Line29 
      BorderStyle     =   2  'Dash
      BorderWidth     =   3
      X1              =   728.376
      X2              =   971.168
      Y1              =   3071.553
      Y2              =   3071.553
   End
   Begin VB.Line Line26 
      BorderWidth     =   3
      Index           =   5
      X1              =   485.584
      X2              =   728.376
      Y1              =   3909.25
      Y2              =   3909.25
   End
   Begin VB.Line Line25 
      BorderWidth     =   3
      Index           =   5
      X1              =   971.168
      X2              =   1213.96
      Y1              =   2233.857
      Y2              =   2233.857
   End
   Begin VB.Line Line24 
      BorderWidth     =   3
      Index           =   5
      X1              =   1213.96
      X2              =   1456.752
      Y1              =   1256.545
      Y2              =   1256.545
   End
   Begin VB.Line Line23 
      BorderWidth     =   3
      Index           =   5
      X1              =   849.772
      X2              =   1092.564
      Y1              =   2652.705
      Y2              =   2652.705
   End
   Begin VB.Line Line22 
      BorderWidth     =   3
      Index           =   5
      X1              =   971.168
      X2              =   1213.96
      Y1              =   1815.009
      Y2              =   1815.009
   End
   Begin VB.Line Line21 
      BorderWidth     =   3
      Index           =   5
      X1              =   1092.564
      X2              =   1335.356
      Y1              =   1535.777
      Y2              =   1535.777
   End
   Begin VB.Line Line20 
      BorderWidth     =   3
      Index           =   5
      X1              =   1456.752
      X2              =   728.376
      Y1              =   977.312
      Y2              =   4048.866
   End
   Begin VB.Line Line19 
      BorderWidth     =   3
      Index           =   5
      X1              =   1213.96
      X2              =   485.584
      Y1              =   977.312
      Y2              =   4048.866
   End
   Begin VB.Line Line28 
      BorderWidth     =   3
      X1              =   1578.148
      X2              =   1942.336
      Y1              =   6841.187
      Y2              =   6841.187
   End
   Begin VB.Line Line27 
      BorderWidth     =   3
      X1              =   1699.544
      X2              =   2063.732
      Y1              =   6980.803
      Y2              =   6980.803
   End
   Begin VB.Line Line18 
      BorderWidth     =   3
      X1              =   1820.94
      X2              =   2185.128
      Y1              =   7120.419
      Y2              =   7120.419
   End
   Begin VB.Line Line26 
      BorderWidth     =   3
      Index           =   2
      X1              =   2549.316
      X2              =   2913.504
      Y1              =   7958.115
      Y2              =   7958.115
   End
   Begin VB.Line Line25 
      BorderWidth     =   3
      Index           =   2
      X1              =   2306.524
      X2              =   2670.712
      Y1              =   7678.883
      Y2              =   7678.883
   End
   Begin VB.Line Line24 
      BorderWidth     =   3
      Index           =   2
      X1              =   1942.336
      X2              =   2306.524
      Y1              =   7260.035
      Y2              =   7260.035
   End
   Begin VB.Line Line23 
      BorderWidth     =   3
      Index           =   2
      X1              =   2427.92
      X2              =   2792.108
      Y1              =   7818.499
      Y2              =   7818.499
   End
   Begin VB.Line Line22 
      BorderWidth     =   3
      Index           =   2
      X1              =   2185.128
      X2              =   2549.316
      Y1              =   7539.267
      Y2              =   7539.267
   End
   Begin VB.Line Line21 
      BorderWidth     =   3
      Index           =   4
      X1              =   2063.732
      X2              =   2427.92
      Y1              =   7399.651
      Y2              =   7399.651
   End
   Begin VB.Line Line20 
      BorderWidth     =   3
      Index           =   2
      X1              =   1699.544
      X2              =   3034.9
      Y1              =   6561.955
      Y2              =   8097.731
   End
   Begin VB.Line Line19 
      BorderWidth     =   3
      Index           =   2
      X1              =   1456.752
      X2              =   2792.108
      Y1              =   6701.571
      Y2              =   8237.348
   End
   Begin VB.Line Line16 
      BorderWidth     =   3
      X1              =   728.376
      X2              =   971.168
      Y1              =   5445.026
      Y2              =   5445.026
   End
   Begin VB.Line Line26 
      BorderWidth     =   3
      Index           =   3
      X1              =   1092.564
      X2              =   1335.356
      Y1              =   7120.419
      Y2              =   7120.419
   End
   Begin VB.Line Line25 
      BorderWidth     =   3
      Index           =   3
      X1              =   971.168
      X2              =   1213.96
      Y1              =   6561.955
      Y2              =   6561.955
   End
   Begin VB.Line Line24 
      BorderWidth     =   3
      Index           =   3
      X1              =   728.376
      X2              =   971.168
      Y1              =   5724.258
      Y2              =   5724.258
   End
   Begin VB.Line Line23 
      BorderWidth     =   3
      Index           =   3
      X1              =   971.168
      X2              =   1213.96
      Y1              =   6841.187
      Y2              =   6841.187
   End
   Begin VB.Line Line22 
      BorderWidth     =   3
      Index           =   3
      X1              =   849.772
      X2              =   1092.564
      Y1              =   6282.723
      Y2              =   6282.723
   End
   Begin VB.Line Line21 
      BorderWidth     =   3
      Index           =   2
      X1              =   849.772
      X2              =   1092.564
      Y1              =   6003.49
      Y2              =   6003.49
   End
   Begin VB.Line Line20 
      BorderWidth     =   3
      Index           =   3
      X1              =   849.772
      X2              =   1335.356
      Y1              =   5165.794
      Y2              =   7260.035
   End
   Begin VB.Line Line19 
      BorderWidth     =   3
      Index           =   3
      X1              =   606.98
      X2              =   1092.564
      Y1              =   5165.794
      Y2              =   7399.651
   End
   Begin VB.Line Line26 
      BorderWidth     =   3
      Index           =   1
      X1              =   6191.196
      X2              =   6433.988
      Y1              =   8097.731
      Y2              =   8097.731
   End
   Begin VB.Line Line25 
      BorderWidth     =   3
      Index           =   1
      X1              =   6191.196
      X2              =   6433.988
      Y1              =   7539.267
      Y2              =   7539.267
   End
   Begin VB.Line Line24 
      BorderWidth     =   3
      Index           =   1
      X1              =   6191.196
      X2              =   6433.988
      Y1              =   6701.571
      Y2              =   6701.571
   End
   Begin VB.Line Line23 
      BorderWidth     =   3
      Index           =   1
      X1              =   6191.196
      X2              =   6433.988
      Y1              =   7818.499
      Y2              =   7818.499
   End
   Begin VB.Line Line22 
      BorderWidth     =   3
      Index           =   1
      X1              =   6191.196
      X2              =   6433.988
      Y1              =   7260.035
      Y2              =   7260.035
   End
   Begin VB.Line Line21 
      BorderWidth     =   3
      Index           =   1
      X1              =   6191.196
      X2              =   6433.988
      Y1              =   6980.803
      Y2              =   6980.803
   End
   Begin VB.Line Line20 
      BorderWidth     =   3
      Index           =   1
      X1              =   6433.988
      X2              =   6433.988
      Y1              =   6561.955
      Y2              =   8237.348
   End
   Begin VB.Line Line19 
      BorderWidth     =   3
      Index           =   1
      X1              =   6191.196
      X2              =   6191.196
      Y1              =   6561.955
      Y2              =   8237.348
   End
   Begin VB.Line Line26 
      BorderWidth     =   3
      Index           =   0
      X1              =   8740.513
      X2              =   8983.305
      Y1              =   6003.49
      Y2              =   6003.49
   End
   Begin VB.Line Line25 
      BorderWidth     =   3
      Index           =   0
      X1              =   8740.513
      X2              =   8983.305
      Y1              =   5445.026
      Y2              =   5445.026
   End
   Begin VB.Line Line24 
      BorderWidth     =   3
      Index           =   0
      X1              =   8740.513
      X2              =   8983.305
      Y1              =   4607.33
      Y2              =   4607.33
   End
   Begin VB.Line Line23 
      BorderWidth     =   3
      Index           =   0
      X1              =   8740.513
      X2              =   8983.305
      Y1              =   5724.258
      Y2              =   5724.258
   End
   Begin VB.Line Line22 
      BorderWidth     =   3
      Index           =   0
      X1              =   8740.513
      X2              =   8983.305
      Y1              =   5165.794
      Y2              =   5165.794
   End
   Begin VB.Line Line21 
      BorderWidth     =   3
      Index           =   0
      X1              =   8740.513
      X2              =   8983.305
      Y1              =   4886.562
      Y2              =   4886.562
   End
   Begin VB.Line Line20 
      BorderWidth     =   3
      Index           =   0
      X1              =   8983.305
      X2              =   8983.305
      Y1              =   4467.714
      Y2              =   6143.106
   End
   Begin VB.Line Line19 
      BorderWidth     =   3
      Index           =   0
      X1              =   8740.513
      X2              =   8740.513
      Y1              =   4467.714
      Y2              =   6143.106
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "End        point"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   99
      Left            =   7800
      TabIndex        =   101
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   98
      Left            =   7200
      TabIndex        =   100
      Top             =   480
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   97
      Left            =   6360
      TabIndex        =   99
      Top             =   480
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   96
      Left            =   5520
      TabIndex        =   98
      Top             =   480
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   95
      Left            =   4680
      TabIndex        =   97
      Top             =   480
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   94
      Left            =   3840
      TabIndex        =   96
      Top             =   480
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   93
      Left            =   3000
      TabIndex        =   95
      Top             =   480
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   92
      Left            =   2160
      TabIndex        =   94
      Top             =   480
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   91
      Left            =   1320
      TabIndex        =   93
      Top             =   480
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   90
      Left            =   480
      TabIndex        =   92
      Top             =   480
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   89
      Left            =   480
      TabIndex        =   91
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   88
      Left            =   1320
      TabIndex        =   90
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   87
      Left            =   2160
      TabIndex        =   89
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   86
      Left            =   3000
      TabIndex        =   88
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   85
      Left            =   3840
      TabIndex        =   87
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   84
      Left            =   4680
      TabIndex        =   86
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   83
      Left            =   5520
      TabIndex        =   85
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   82
      Left            =   6360
      TabIndex        =   84
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   81
      Left            =   7200
      TabIndex        =   83
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   80
      Left            =   8040
      TabIndex        =   82
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   79
      Left            =   8040
      TabIndex        =   81
      Top             =   1920
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   78
      Left            =   7200
      TabIndex        =   80
      Top             =   1920
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   77
      Left            =   6360
      TabIndex        =   79
      Top             =   1920
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   76
      Left            =   5520
      TabIndex        =   78
      Top             =   1920
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   75
      Left            =   4680
      TabIndex        =   77
      Top             =   1920
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   74
      Left            =   3840
      TabIndex        =   76
      Top             =   1920
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   73
      Left            =   3000
      TabIndex        =   75
      Top             =   1920
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   72
      Left            =   2160
      TabIndex        =   74
      Top             =   1920
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   71
      Left            =   1320
      TabIndex        =   73
      Top             =   1920
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   70
      Left            =   480
      TabIndex        =   72
      Top             =   1920
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   69
      Left            =   480
      TabIndex        =   71
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   68
      Left            =   1320
      TabIndex        =   70
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   67
      Left            =   2160
      TabIndex        =   69
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   65
      Left            =   3840
      TabIndex        =   68
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   64
      Left            =   4680
      TabIndex        =   67
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   63
      Left            =   5520
      TabIndex        =   66
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   62
      Left            =   6360
      TabIndex        =   65
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   61
      Left            =   7200
      TabIndex        =   64
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   60
      Left            =   8040
      TabIndex        =   63
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   59
      Left            =   8040
      TabIndex        =   62
      Top             =   3360
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   58
      Left            =   7200
      TabIndex        =   61
      Top             =   3360
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   57
      Left            =   6360
      TabIndex        =   60
      Top             =   3360
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   56
      Left            =   5520
      TabIndex        =   59
      Top             =   3360
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   55
      Left            =   4680
      TabIndex        =   58
      Top             =   3360
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   53
      Left            =   3000
      TabIndex        =   57
      Top             =   3360
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   52
      Left            =   2160
      TabIndex        =   56
      Top             =   3360
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   51
      Left            =   1320
      TabIndex        =   55
      Top             =   3360
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   50
      Left            =   480
      TabIndex        =   54
      Top             =   3360
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   49
      Left            =   480
      TabIndex        =   53
      Top             =   4080
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   48
      Left            =   1320
      TabIndex        =   52
      Top             =   4080
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   47
      Left            =   2160
      TabIndex        =   51
      Top             =   4080
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   46
      Left            =   2760
      TabIndex        =   50
      Top             =   4080
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   45
      Left            =   3840
      TabIndex        =   49
      Top             =   4080
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   44
      Left            =   4680
      TabIndex        =   48
      Top             =   4080
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   43
      Left            =   5520
      TabIndex        =   47
      Top             =   4080
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   42
      Left            =   6360
      TabIndex        =   46
      Top             =   4080
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   41
      Left            =   7200
      TabIndex        =   45
      Top             =   4080
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   40
      Left            =   8040
      TabIndex        =   44
      Top             =   4080
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   39
      Left            =   8040
      TabIndex        =   43
      Top             =   4800
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   38
      Left            =   7200
      TabIndex        =   42
      Top             =   4800
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   37
      Left            =   6360
      TabIndex        =   41
      Top             =   4800
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   36
      Left            =   5520
      TabIndex        =   40
      Top             =   4800
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   35
      Left            =   4680
      TabIndex        =   39
      Top             =   4800
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   34
      Left            =   3840
      TabIndex        =   38
      Top             =   4800
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   33
      Left            =   3000
      TabIndex        =   37
      Top             =   4800
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   32
      Left            =   2160
      TabIndex        =   36
      Top             =   4800
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   31
      Left            =   1320
      TabIndex        =   35
      Top             =   4800
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   30
      Left            =   480
      TabIndex        =   34
      Top             =   4800
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   29
      Left            =   480
      TabIndex        =   33
      Top             =   5520
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   28
      Left            =   1320
      TabIndex        =   32
      Top             =   5520
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   27
      Left            =   2160
      TabIndex        =   31
      Top             =   5520
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   26
      Left            =   3000
      TabIndex        =   30
      Top             =   5520
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   25
      Left            =   3840
      TabIndex        =   29
      Top             =   5520
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   24
      Left            =   4680
      TabIndex        =   28
      Top             =   5520
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   23
      Left            =   5520
      TabIndex        =   27
      Top             =   5520
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   22
      Left            =   6360
      TabIndex        =   26
      Top             =   5520
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   21
      Left            =   7200
      TabIndex        =   25
      Top             =   5520
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   20
      Left            =   8040
      TabIndex        =   24
      Top             =   5520
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   19
      Left            =   8040
      TabIndex        =   23
      Top             =   6240
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   18
      Left            =   7200
      TabIndex        =   22
      Top             =   6240
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   17
      Left            =   6360
      TabIndex        =   21
      Top             =   6240
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   16
      Left            =   5520
      TabIndex        =   20
      Top             =   6240
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   15
      Left            =   4680
      TabIndex        =   19
      Top             =   6240
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   14
      Left            =   3840
      TabIndex        =   18
      Top             =   6240
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   13
      Left            =   3000
      TabIndex        =   17
      Top             =   6240
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   12
      Left            =   2160
      TabIndex        =   16
      Top             =   6240
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   11
      Left            =   1320
      TabIndex        =   15
      Top             =   6240
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   10
      Left            =   480
      TabIndex        =   14
      Top             =   6240
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   480
      TabIndex        =   13
      Top             =   6960
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   1320
      TabIndex        =   12
      Top             =   6960
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   2160
      TabIndex        =   11
      Top             =   6960
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   3000
      TabIndex        =   10
      Top             =   6960
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
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
      Height          =   375
      Index           =   5
      Left            =   3840
      TabIndex        =   9
      Top             =   6960
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
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
      Height          =   375
      Index           =   4
      Left            =   4800
      TabIndex        =   8
      Top             =   7080
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   5520
      TabIndex        =   7
      Top             =   6960
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   6360
      TabIndex        =   6
      Top             =   6960
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   7200
      TabIndex        =   5
      Top             =   6960
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Start        Point"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   7800
      TabIndex        =   4
      Top             =   6960
      Width           =   735
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Index           =   35
      Left            =   4320
      Top             =   4560
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Index           =   49
      Left            =   120
      Top             =   3840
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Index           =   99
      Left            =   7680
      Top             =   240
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Index           =   98
      Left            =   6840
      Top             =   240
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Index           =   97
      Left            =   6000
      Top             =   240
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Index           =   96
      Left            =   5160
      Top             =   240
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Index           =   95
      Left            =   4320
      Top             =   240
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Index           =   94
      Left            =   3480
      Top             =   240
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Index           =   93
      Left            =   2640
      Top             =   240
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Index           =   92
      Left            =   1800
      Top             =   240
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Index           =   91
      Left            =   960
      Top             =   240
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Index           =   90
      Left            =   120
      Top             =   240
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Index           =   89
      Left            =   120
      Top             =   960
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Index           =   88
      Left            =   960
      Top             =   960
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Index           =   87
      Left            =   1800
      Top             =   960
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Index           =   86
      Left            =   2640
      Top             =   960
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Index           =   85
      Left            =   3480
      Top             =   960
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Index           =   84
      Left            =   4320
      Top             =   960
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Index           =   83
      Left            =   5160
      Top             =   960
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Index           =   82
      Left            =   6000
      Top             =   960
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Index           =   81
      Left            =   6840
      Top             =   960
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Index           =   80
      Left            =   7680
      Top             =   960
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Index           =   79
      Left            =   7680
      Top             =   1680
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Index           =   78
      Left            =   6840
      Top             =   1680
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Index           =   77
      Left            =   6000
      Top             =   1680
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Index           =   76
      Left            =   5160
      Top             =   1680
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Index           =   75
      Left            =   4320
      Top             =   1680
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Index           =   74
      Left            =   3480
      Top             =   1680
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Index           =   73
      Left            =   2640
      Top             =   1680
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Index           =   72
      Left            =   1800
      Top             =   1680
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Index           =   71
      Left            =   960
      Top             =   1680
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Index           =   70
      Left            =   120
      Top             =   1680
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Index           =   69
      Left            =   120
      Top             =   2400
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Index           =   68
      Left            =   960
      Top             =   2400
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Index           =   67
      Left            =   1800
      Top             =   2400
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Index           =   66
      Left            =   2640
      Top             =   2400
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Index           =   65
      Left            =   3480
      Top             =   2400
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Index           =   64
      Left            =   4320
      Top             =   2400
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Index           =   63
      Left            =   5160
      Top             =   2400
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Index           =   62
      Left            =   6000
      Top             =   2400
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Index           =   61
      Left            =   6840
      Top             =   2400
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Index           =   60
      Left            =   7680
      Top             =   2400
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Index           =   58
      Left            =   6840
      Top             =   3120
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Index           =   57
      Left            =   6000
      Top             =   3120
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Index           =   56
      Left            =   5160
      Top             =   3120
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Index           =   55
      Left            =   4320
      Top             =   3120
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Index           =   54
      Left            =   3480
      Top             =   3120
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Index           =   53
      Left            =   2640
      Top             =   3120
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Index           =   52
      Left            =   1800
      Top             =   3120
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Index           =   51
      Left            =   960
      Top             =   3120
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Index           =   50
      Left            =   120
      Top             =   3120
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Index           =   48
      Left            =   960
      Top             =   3840
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Index           =   47
      Left            =   1800
      Top             =   3840
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Index           =   46
      Left            =   2640
      Top             =   3840
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Index           =   45
      Left            =   3480
      Top             =   3840
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Index           =   44
      Left            =   4320
      Top             =   3840
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Index           =   43
      Left            =   5160
      Top             =   3840
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Index           =   42
      Left            =   6000
      Top             =   3840
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Index           =   41
      Left            =   6840
      Top             =   3840
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Index           =   40
      Left            =   7680
      Top             =   3840
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Index           =   39
      Left            =   7680
      Top             =   4560
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Index           =   38
      Left            =   6840
      Top             =   4560
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Index           =   37
      Left            =   6000
      Top             =   4560
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Index           =   36
      Left            =   5160
      Top             =   4560
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Index           =   34
      Left            =   3480
      Top             =   4560
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Index           =   33
      Left            =   2640
      Top             =   4560
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Index           =   32
      Left            =   1800
      Top             =   4560
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Index           =   31
      Left            =   960
      Top             =   4560
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Index           =   30
      Left            =   120
      Top             =   4560
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Index           =   29
      Left            =   120
      Top             =   5280
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Index           =   28
      Left            =   960
      Top             =   5280
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Index           =   27
      Left            =   1800
      Top             =   5280
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Index           =   26
      Left            =   2640
      Top             =   5280
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Index           =   25
      Left            =   3480
      Top             =   5280
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Index           =   24
      Left            =   4320
      Top             =   5280
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Index           =   23
      Left            =   5160
      Top             =   5280
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Index           =   22
      Left            =   6000
      Top             =   5280
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Index           =   21
      Left            =   6840
      Top             =   5280
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Index           =   20
      Left            =   7680
      Top             =   5280
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Index           =   19
      Left            =   7680
      Top             =   6000
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Index           =   18
      Left            =   6840
      Top             =   6000
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Index           =   17
      Left            =   6000
      Top             =   6000
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Index           =   16
      Left            =   5160
      Top             =   6000
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Index           =   15
      Left            =   4320
      Top             =   6000
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Index           =   14
      Left            =   3480
      Top             =   6000
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Index           =   13
      Left            =   2640
      Top             =   6000
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Index           =   12
      Left            =   1800
      Top             =   6000
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Index           =   11
      Left            =   960
      Top             =   6000
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Index           =   10
      Left            =   120
      Top             =   6000
      Width           =   855
   End
   Begin VB.Image Image2 
      Height          =   405
      Left            =   8640
      Picture         =   "game.frx":098A
      Top             =   7080
      Width           =   435
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Index           =   9
      Left            =   120
      Top             =   6720
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Index           =   8
      Left            =   960
      Top             =   6720
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Index           =   7
      Left            =   1800
      Top             =   6720
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Index           =   6
      Left            =   2640
      Top             =   6720
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Index           =   5
      Left            =   3480
      Top             =   6720
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Index           =   4
      Left            =   4320
      Top             =   6720
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Index           =   3
      Left            =   5160
      Top             =   6720
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Index           =   2
      Left            =   6000
      Top             =   6720
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Index           =   1
      Left            =   6840
      Top             =   6720
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Index           =   0
      Left            =   7680
      Top             =   6720
      Width           =   855
   End
   Begin VB.Line Line9 
      X1              =   7769.344
      X2              =   7769.344
      Y1              =   8656.195
      Y2              =   279.232
   End
   Begin VB.Line Line8 
      X1              =   6919.572
      X2              =   6919.572
      Y1              =   8656.195
      Y2              =   279.232
   End
   Begin VB.Line Line7 
      X1              =   6069.8
      X2              =   6069.8
      Y1              =   8656.195
      Y2              =   279.232
   End
   Begin VB.Line Line6 
      X1              =   5220.028
      X2              =   5220.028
      Y1              =   8656.195
      Y2              =   279.232
   End
   Begin VB.Line Line5 
      X1              =   4370.256
      X2              =   4370.256
      Y1              =   8656.195
      Y2              =   279.232
   End
   Begin VB.Line Line4 
      X1              =   3520.484
      X2              =   3520.484
      Y1              =   8656.195
      Y2              =   279.232
   End
   Begin VB.Line Line3 
      X1              =   2670.712
      X2              =   2670.712
      Y1              =   8656.195
      Y2              =   279.232
   End
   Begin VB.Line Line2 
      X1              =   1820.94
      X2              =   1820.94
      Y1              =   8656.195
      Y2              =   279.232
   End
   Begin VB.Line Line1 
      X1              =   971.168
      X2              =   971.168
      Y1              =   279.232
      Y2              =   8656.195
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   99
      Left            =   7680
      Top             =   240
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   98
      Left            =   6840
      Top             =   240
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   96
      Left            =   5160
      Top             =   240
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   94
      Left            =   3480
      Top             =   240
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   91
      Left            =   960
      Top             =   240
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   90
      Left            =   120
      Top             =   240
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   89
      Left            =   120
      Top             =   960
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   87
      Left            =   1800
      Top             =   960
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   82
      Left            =   6000
      Top             =   960
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   81
      Left            =   6840
      Top             =   960
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   80
      Left            =   7680
      Top             =   960
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   76
      Left            =   5160
      Top             =   1680
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   74
      Left            =   3480
      Top             =   1680
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   73
      Left            =   2640
      Top             =   1680
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   72
      Left            =   1800
      Top             =   1680
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   71
      Left            =   960
      Top             =   1680
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   70
      Left            =   120
      Top             =   1680
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   63
      Left            =   5160
      Top             =   2400
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   62
      Left            =   6000
      Top             =   2400
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   61
      Left            =   6840
      Top             =   2400
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   54
      Left            =   3480
      Top             =   3120
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   53
      Left            =   2640
      Top             =   3120
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   52
      Left            =   1800
      Top             =   3120
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   51
      Left            =   960
      Top             =   3120
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   48
      Left            =   960
      Top             =   3840
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   46
      Left            =   3480
      Top             =   3840
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   44
      Left            =   4320
      Top             =   3840
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   41
      Left            =   6840
      Top             =   3840
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   40
      Left            =   7680
      Top             =   3840
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   36
      Left            =   5160
      Top             =   4560
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   30
      Left            =   120
      Top             =   4560
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   28
      Left            =   960
      Top             =   5280
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   26
      Left            =   2640
      Top             =   5280
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   24
      Left            =   4320
      Top             =   5280
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   23
      Left            =   5160
      Top             =   5280
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   17
      Left            =   6000
      Top             =   6000
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   16
      Left            =   5160
      Top             =   6000
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   13
      Left            =   2640
      Top             =   6000
      Width           =   855
   End
   Begin VB.Shape Shape2 
      Height          =   735
      Index           =   9
      Left            =   120
      Top             =   240
      Width           =   8415
   End
   Begin VB.Shape Shape2 
      Height          =   735
      Index           =   8
      Left            =   120
      Top             =   960
      Width           =   8415
   End
   Begin VB.Shape Shape2 
      Height          =   735
      Index           =   7
      Left            =   120
      Top             =   1680
      Width           =   8415
   End
   Begin VB.Shape Shape2 
      Height          =   735
      Index           =   6
      Left            =   120
      Top             =   2400
      Width           =   8415
   End
   Begin VB.Shape Shape2 
      Height          =   735
      Index           =   5
      Left            =   120
      Top             =   3120
      Width           =   8415
   End
   Begin VB.Shape Shape2 
      Height          =   735
      Index           =   4
      Left            =   120
      Top             =   3840
      Width           =   8415
   End
   Begin VB.Shape Shape2 
      Height          =   735
      Index           =   3
      Left            =   120
      Top             =   4560
      Width           =   8415
   End
   Begin VB.Shape Shape2 
      Height          =   735
      Index           =   2
      Left            =   120
      Top             =   5280
      Width           =   8415
   End
   Begin VB.Shape Shape2 
      Height          =   735
      Index           =   1
      Left            =   120
      Top             =   6000
      Width           =   8415
   End
   Begin VB.Shape Shape2 
      Height          =   735
      Index           =   0
      Left            =   120
      Top             =   6720
      Width           =   8415
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      Height          =   7215
      Left            =   120
      Top             =   240
      Width           =   8415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim prev, no, b(100), a(100), i, k, c, h, cm, hm, outh, outc
Dim w As Label
Private Sub Command1_Click()
q: Randomize
no = Int((6 * Rnd) + 1)
If no = prev Then
GoTo q
End If
ShockwaveFlash1.GotoFrame (no - 1)
prev = no
Command2.Enabled = True
Command1.Enabled = False
End Sub

Private Sub Command2_Click()
hm = hm + 1
If k + no >= 99 And outh = 1 Then
Image2.Move b(99), a(99)
MsgBox "Game Over,You Win in " & hm & " Moves"
Set rs = New Recordset
Set con = New Connection
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\updown.mdb;Persist Security Info=False"
rs.Open "select * from updown", con, adOpenKeyset, adLockOptimistic
rs.AddNew Array("Name", "moves"), Array(Form3.Text1.Text, hm)
ElseIf outh = 1 And k + no < 99 Then
Select Case k + no + 1
Case 3: k = 23 - 1
Case 7: k = 29 - 1
Case 12: k = 50 - 1
Case 17: k = 37 - 1
Case 21: k = 43 - 1
Case 33: k = 67 - 1
Case 41: k = 61 - 1
Case 48: k = 68 - 1
Case 51: k = 92 - 1
Case 57: k = 63 - 1
Case 62: k = 82 - 1
Case 77: k = 97 - 1
Case 86: k = 95 - 1

Case 99: k = 56 - 1
Case 93: k = 49 - 1
Case 88: k = 66 - 1
Case 78: k = 58 - 1
Case 31: k = 10 - 1
Case 55: k = 6 - 1
Case 57: k = 26 - 1
Case 42: k = 19 - 1
Case 33: k = 9 - 1
Case Else: k = k + no
End Select

Image2.Move b(k), a(k)
ElseIf no = 6 And outh = 0 Then
outh = 1
Image2.Move b(k), a(k)
End If
Text1.Text = "Computer To Play"
Form1.Refresh
Delay (2)
Call Command1_Click
Call comp
End Sub

Private Sub Command3_Click()
Form4.Show
End Sub

Private Sub Command4_Click()
End
End Sub

Private Sub Form_Load()
Command2.Enabled = False
prev = 0
For i = 0 To 99
a(i) = Shape3(i).Top + 60
b(i) = Shape3(i).Left + 70
Next
i = 0
For i = 1 To 99
Label1(i).FontBold = True
Label1(i).Caption = i + 1
Next
k = 0
c = 0
outh = 0
outc = 0
Text1.Text = Form3.Text1.Text & " To Play"
End Sub


Public Sub comp()
cm = cm + 1
If c + no >= 99 And outc = 1 Then
Image3.Move b(99), a(99)
MsgBox "Game Over,Computer Wins in " & cm & " Moves"
Set rs = New Recordset
Set con = New Connection
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\updown.mdb;Persist Security Info=False"
rs.Open "select * from updown", con, adOpenKeyset, adLockOptimistic
rs.AddNew Array("name", "moves"), Array("Computer", cm)
ElseIf outc = 1 And c + no < 99 Then
Select Case c + no + 1
Case 3: c = 23 - 1
Case 7: c = 29 - 1
Case 12: c = 50 - 1
Case 17: c = 37 - 1
Case 21: c = 43 - 1
Case 33: c = 67 - 1
Case 41: c = 61 - 1
Case 48: c = 68 - 1
Case 51: c = 92 - 1
Case 57: c = 63 - 1
Case 62: c = 82 - 1
Case 77: c = 97 - 1
Case 86: c = 95 - 1

Case 99: c = 56 - 1
Case 93: c = 49 - 1
Case 88: c = 66 - 1
Case 78: c = 58 - 1
Case 31: c = 10 - 1
Case 55: c = 6 - 1
Case 57: c = 26 - 1
Case 42: c = 19 - 1
Case 33: c = 9 - 1
Case Else: c = c + no
End Select

Image3.Move b(c), a(c)
ElseIf no = 6 And outc = 0 Then
outc = 1
Image3.Move b(c), a(c)
End If
Command2.Enabled = False
Command1.Enabled = True
Text1.Text = Form3.Text1.Text & " To Play"
End Sub
