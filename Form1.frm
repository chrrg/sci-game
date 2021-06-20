VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000013&
   BorderStyle     =   0  'None
   ClientHeight    =   9270
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11805
   FillColor       =   &H80000012&
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9270
   ScaleWidth      =   11805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   WindowState     =   2  'Maximized
   Begin VB.PictureBox didi2 
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   4440
      ScaleHeight     =   375
      ScaleWidth      =   3735
      TabIndex        =   42
      Top             =   2400
      Width           =   3735
   End
   Begin VB.Frame Frame1 
      Caption         =   "作弊器"
      Height          =   4815
      Left            =   600
      TabIndex        =   25
      Top             =   1920
      Width           =   2655
      Begin VB.CommandButton Command16 
         Caption         =   "Command16"
         Height          =   375
         Left            =   1320
         TabIndex        =   46
         Top             =   3600
         Width           =   975
      End
      Begin VB.CommandButton Command8 
         Caption         =   "经验+"
         Height          =   495
         Left            =   1320
         TabIndex        =   44
         Top             =   3120
         Width           =   975
      End
      Begin VB.CommandButton Command15 
         Caption         =   "经验+"
         Height          =   495
         Left            =   360
         TabIndex        =   45
         Top             =   3120
         Width           =   975
      End
      Begin VB.CommandButton Command7 
         Caption         =   "-"
         Height          =   255
         Left            =   2280
         TabIndex        =   40
         Top             =   0
         Width           =   375
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H80000001&
         Height          =   375
         Left            =   360
         TabIndex        =   39
         Top             =   3960
         Width           =   1935
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000001&
         Height          =   375
         Left            =   360
         TabIndex        =   38
         Top             =   4320
         Width           =   1935
      End
      Begin VB.CommandButton Command11 
         Caption         =   "自残"
         Height          =   495
         Left            =   1320
         TabIndex        =   30
         Top             =   2640
         Width           =   975
      End
      Begin VB.CommandButton Command13 
         Caption         =   "自残"
         Height          =   495
         Left            =   360
         TabIndex        =   28
         Top             =   2640
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "攻击"
         Height          =   495
         Left            =   1320
         TabIndex        =   37
         Top             =   2160
         Width           =   975
      End
      Begin VB.CommandButton Command12 
         Caption         =   "攻击"
         Height          =   495
         Left            =   360
         TabIndex        =   29
         Top             =   2160
         Width           =   975
      End
      Begin VB.CommandButton Command14 
         Caption         =   "HP"
         Height          =   495
         Left            =   1320
         TabIndex        =   27
         Top             =   1680
         Width           =   975
      End
      Begin VB.CommandButton Command9 
         Caption         =   "X"
         Height          =   495
         Left            =   360
         TabIndex        =   32
         Top             =   1680
         Width           =   975
      End
      Begin VB.CommandButton Command3 
         Caption         =   "敌方增兵"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1320
         TabIndex        =   35
         Top             =   1200
         Width           =   975
      End
      Begin VB.CommandButton Command6 
         Caption         =   "对方死"
         Height          =   495
         Left            =   1320
         TabIndex        =   33
         Top             =   720
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "我方增兵"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   36
         Top             =   1200
         Width           =   975
      End
      Begin VB.CommandButton Command4 
         Caption         =   "时间"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1320
         TabIndex        =   34
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command10 
         Caption         =   "自己死"
         Height          =   495
         Left            =   360
         TabIndex        =   31
         Top             =   720
         Width           =   975
      End
      Begin VB.CommandButton Command17 
         Caption         =   "更新视野"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   26
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.PictureBox rpic 
      Height          =   615
      Left            =   4200
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   24
      Top             =   1200
      Width           =   735
   End
   Begin VB.PictureBox epic 
      Height          =   615
      Left            =   3960
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   23
      Top             =   1200
      Width           =   735
   End
   Begin VB.PictureBox wpic 
      Height          =   615
      Left            =   3720
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   22
      Top             =   1200
      Width           =   735
   End
   Begin VB.PictureBox qpic 
      Height          =   615
      Left            =   3480
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   21
      Top             =   1200
      Width           =   735
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H8000000D&
      Height          =   375
      Left            =   600
      TabIndex        =   20
      Text            =   "64"
      Top             =   360
      Width           =   1575
   End
   Begin VB.Timer keys 
      Interval        =   10
      Left            =   8640
      Top             =   4320
   End
   Begin VB.CommandButton ends 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11400
      TabIndex        =   6
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton abouts 
      Caption         =   "关于"
      Height          =   375
      Left            =   10680
      TabIndex        =   19
      Top             =   0
      Width           =   735
   End
   Begin VB.Timer over 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   8640
      Top             =   5400
   End
   Begin VB.PictureBox bar 
      Height          =   495
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   11715
      TabIndex        =   17
      Top             =   7800
      Width           =   11775
   End
   Begin VB.PictureBox Bpic 
      Height          =   7815
      Left            =   0
      ScaleHeight     =   7755
      ScaleWidth      =   195
      TabIndex        =   16
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Timer dieds2 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   8640
      Top             =   5040
   End
   Begin VB.Timer dieds 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   8640
      Top             =   4680
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H8000000D&
      Caption         =   "继续游戏"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   16.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      MaskColor       =   &H00C0E0FF&
      TabIndex        =   14
      Top             =   4440
      UseMaskColor    =   -1  'True
      Width           =   2175
   End
   Begin VB.PictureBox ziji 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5640
      ScaleHeight     =   795
      ScaleWidth      =   75
      TabIndex        =   12
      Top             =   1080
      Width           =   135
   End
   Begin VB.Timer Timer2 
      Interval        =   200
      Left            =   8640
      Top             =   3960
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   8640
      Top             =   3600
   End
   Begin VB.PictureBox MPT 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   11715
      TabIndex        =   10
      Top             =   8760
      Width           =   11775
   End
   Begin VB.PictureBox HPT 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   11715
      TabIndex        =   9
      Top             =   8280
      Width           =   11775
   End
   Begin VB.PictureBox di 
      Height          =   855
      Left            =   6720
      ScaleHeight     =   795
      ScaleWidth      =   75
      TabIndex        =   15
      Top             =   1080
      Width           =   135
   End
   Begin VB.PictureBox 基地2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6600
      ScaleHeight     =   795
      ScaleWidth      =   75
      TabIndex        =   8
      Top             =   1080
      Width           =   135
   End
   Begin VB.PictureBox 基地1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5760
      ScaleHeight     =   795
      ScaleWidth      =   75
      TabIndex        =   7
      Top             =   1080
      Width           =   135
   End
   Begin VB.PictureBox 水晶2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6480
      ScaleHeight     =   795
      ScaleWidth      =   75
      TabIndex        =   5
      Top             =   1080
      Width           =   135
   End
   Begin VB.PictureBox 塔4 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6360
      ScaleHeight     =   795
      ScaleWidth      =   75
      TabIndex        =   4
      Top             =   1080
      Width           =   135
   End
   Begin VB.PictureBox 塔3 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6240
      ScaleHeight     =   795
      ScaleWidth      =   75
      TabIndex        =   3
      Top             =   1080
      Width           =   135
   End
   Begin VB.PictureBox 塔2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6120
      ScaleHeight     =   795
      ScaleWidth      =   75
      TabIndex        =   2
      Top             =   1080
      Width           =   135
   End
   Begin VB.PictureBox 塔1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6000
      ScaleHeight     =   795
      ScaleWidth      =   75
      TabIndex        =   1
      Top             =   1080
      Width           =   135
   End
   Begin VB.PictureBox 水晶1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5880
      ScaleHeight     =   795
      ScaleWidth      =   75
      TabIndex        =   0
      Top             =   1080
      Width           =   135
   End
   Begin VB.CommandButton zuobi 
      Caption         =   "作弊器"
      Height          =   375
      Left            =   9720
      TabIndex        =   41
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton Command18 
      Caption         =   "开流畅"
      Height          =   375
      Left            =   8640
      TabIndex        =   47
      Top             =   0
      Width           =   1095
   End
   Begin VB.Shape wp2 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   735
      Left            =   8280
      Shape           =   4  'Rounded Rectangle
      Top             =   1200
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label didi 
      BackStyle       =   0  'Transparent
      Caption         =   "游戏开始"
      ForeColor       =   &H00808000&
      Height          =   255
      Left            =   4440
      TabIndex        =   43
      Top             =   2160
      Width           =   3735
   End
   Begin VB.Shape wp 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   735
      Left            =   8400
      Shape           =   4  'Rounded Rectangle
      Top             =   2520
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "          赛"
      Height          =   210
      Left            =   5280
      TabIndex        =   18
      Top             =   3240
      Width           =   1260
   End
   Begin VB.Line L4 
      BorderColor     =   &H000000FF&
      Visible         =   0   'False
      X1              =   4440
      X2              =   6840
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line L3 
      BorderColor     =   &H000000FF&
      Visible         =   0   'False
      X1              =   4440
      X2              =   6840
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line L2 
      BorderColor     =   &H000000FF&
      Visible         =   0   'False
      X1              =   4440
      X2              =   6840
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line L1 
      BorderColor     =   &H000000FF&
      Visible         =   0   'False
      X1              =   4440
      X2              =   6840
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label fpsshow 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FPS:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   600
      TabIndex        =   13
      Top             =   0
      Width           =   720
   End
   Begin VB.Label infs 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "赛"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   840
      Left            =   5880
      TabIndex        =   11
      Top             =   3480
      Width           =   840
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function SetWindowPos& Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Const WS_EX_LAYERED = &H80000
Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA = &H2
Private Const LWA_COLORKEY = &H1
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Type POINTAPI
  X As Long
  Y As Long
End Type
'Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Dim pcback As String
Dim p As POINTAPI
Dim fpss, fps, fpss2
Dim wantx, want2x
Dim rehp
Dim TotalDied, TotalDied2
Dim exp1 As Double
Dim exp2 As Double
Dim jian1 As Picture, jian2 As Picture
Dim pic1 As Currency
Dim tov As Boolean
Dim oldxx
Dim canq1, canw1, cane1, canr1
Dim canq2, canw2, cane2, canr2
Const 基地高 = 100
Const 水晶高 = 2070
Const 塔高 = 3700
Const 英雄高 = 1000
Const 小兵高 = 500
Const Y = 6730
Dim Bs As Long, ox
Dim Tpic
Dim 基地1X
Dim 水晶1X
Dim 塔1X
Dim 塔2X
Dim 塔3X
Dim 塔4X
Dim 水晶2X
Dim 基地2X
Dim b1(20) As PictureBox
Dim b2(20) As PictureBox
Dim daodan(10) As PictureBox
Dim oinf, TotalHP1, TotalHP2, times As Double, reb, diet, hp, MP, Maxhp, Maxmp, X, diX, key, cana, cana2, fang
Dim fpscpu As Boolean
Private Type B
e As Boolean
hp As Long
X As Long
want As Long
f As Long ' 0 left 1 right
ty As Double
act As Long
End Type
Dim b11(20) As B
Dim b22(20) As B
Dim downv As Long, downvme As Long
Dim daodan0(10) As B
Dim T1hp, T2hp, T3hp, T4hp, diHP, diMP, diMaxHP, diMaxMP
Dim xa1, xb1, xc1
Dim bingid
Sub ditox(X)

want2x = X
End Sub
Function AppPath() As String '无\
AppPath = App.Path
If Right(AppPath, 1) = "\" Then AppPath = Left(AppPath, Len(AppPath) - 1)
End Function
Function sTot(tt As Long) As String
  '把秒数转换为时间
  Dim i As Long, h As Integer, m As Integer, s As Integer
  i = tt Mod 86400  '24小时的秒数为86400秒，必须过滤掉才是正确的一天的时间
  h = Int(i / 3600) '算出小时
  m = Int((i Mod 3600) / 60) '算出分钟
  s = (i Mod 3600) Mod 60 '算出秒数
  sTot = Trim(h) & ":" & Trim(m) & ":" & Trim(s) '组合成正确的时间格式
End Function
Sub l10(x1, y1, x2, y2)
L1.Visible = True
L1.x1 = x1
L1.y1 = y1
L1.x2 = x2
L1.y2 = y2
End Sub
Function lev1()
lev1 = 1
Select Case exp1
Case Is > 450
lev1 = 18
Case Is > 400
lev1 = 17
Case Is > 375
lev1 = 16
Case Is > 330
lev1 = 15
Case Is > 300
lev1 = 14
Case Is > 275
lev1 = 13
Case Is > 250
lev1 = 12
Case Is > 225
lev1 = 11
Case Is > 200
lev1 = 10
Case Is > 175
lev1 = 9
Case Is > 150
lev1 = 8
Case Is > 135
lev1 = 7
Case Is > 110
lev1 = 6
Case Is > 90
lev1 = 5
Case Is > 65
lev1 = 4
Case Is > 40
lev1 = 3
Case Is > 20
lev1 = 2
End Select
End Function
Function lev2()
lev2 = 1
Select Case exp2
Case Is > 450
lev2 = 18
Case Is > 400
lev2 = 17
Case Is > 375
lev2 = 16
Case Is > 330
lev2 = 15
Case Is > 300
lev2 = 14
Case Is > 275
lev2 = 13
Case Is > 250
lev2 = 12
Case Is > 225
lev2 = 11
Case Is > 200
lev2 = 10
Case Is > 175
lev2 = 9
Case Is > 150
lev2 = 8
Case Is > 135
lev2 = 7
Case Is > 110
lev2 = 6
Case Is > 90
lev2 = 5
Case Is > 65
lev2 = 4
Case Is > 40
lev2 = 3
Case Is > 20
lev2 = 2
End Select
End Function
Sub updateview()
If Rnd < 0.5 Then Exit Sub
Me.Cls
    xs = X - (Me.Width - ziji.Width) / 2
    基地1.Left = 基地1X - xs
    水晶1.Left = 水晶1X - xs
    塔1.Left = 塔1X - xs
    塔2.Left = 塔2X - xs
    塔3.Left = 塔3X - xs
    塔4.Left = 塔4X - xs
    水晶2.Left = 水晶2X - xs
    基地2.Left = 基地2X - xs
    wp2.Left = diX - xs - 5000
    
    di.Left = diX - xs
    ziji.Left = X - xs
    For a = 1 To 20
        If b11(a).e Then
            b1(a).Left = b11(a).X - xs
        End If
        If b22(a).e Then
            b2(a).Left = b22(a).X - xs
        End If
        If a <= 10 Then
            If daodan0(a).e Then
                daodan(a).Left = daodan0(a).X - xs
            End If
        End If
    Next
End Sub
Function isvb6() As Boolean
On Error GoTo d:
Debug.Print 0 / 0
isvb6 = False
Exit Function
d:
isvb6 = True
End Function
Sub l20(x1, y1, x2, y2)
L2.Visible = True
L2.x1 = x1
L2.y1 = y1
L2.x2 = x2
L2.y2 = y2
End Sub
Sub l30(x1, y1, x2, y2)
L3.Visible = True
L3.x1 = x1
L3.y1 = y1
L3.x2 = x2
L3.y2 = y2
End Sub
Sub l40(x1, y1, x2, y2)
L4.Visible = True
L4.x1 = x1
L4.y1 = y1
L4.x2 = x2
L4.y2 = y2
End Sub
Sub dd(xx, wx, t)
For a = 1 To 10
If daodan0(a).e = False Then
daodan0(a).e = True
daodan0(a).X = xx + ziji.Left
daodan0(a).want = wx + ziji.Left
daodan0(a).ty = t
daodan(a).Left = -1500
daodan(a).Visible = True
If wx > xx Then
daodan0(a).f = 1
daodan(a).Picture = jian2
Else
daodan0(a).f = 0
daodan(a).Picture = jian1
End If
Exit For
End If
Next
End Sub
Sub c1(hps)
hp = hps
If hp <= 0 Then inf "你已经死了！等待重生.":  died: Exit Sub
If hp > Maxhp Then hp = Maxhp
HPT.Width = Me.Width / 2 * hp / Maxhp

End Sub
Function c2(mps) As Boolean
If mps < 0 Then inf "无蓝": c2 = False: Exit Function
MP = mps
If MP > Maxmp Then MP = Maxmp
MPT.Width = Me.Width / 2 * MP / Maxmp
c2 = True
End Function
Sub pr(s)
Me.Cls
Print ""
Print ""
Print ""
Print ""
Print s
End Sub
Sub died()
On Error Resume Next
c1 (1)
c2 (1)
ziji.Visible = False
infs.Caption = "你已经死了!"
X = 0
'hp = 0
wantx = ""
updateview
dieds.Interval = 200000
If TotalDied < 19 Then
dieds.Interval = 1000 * TotalDied * 2 + 10000
Else
dieds.Interval = 500000
End If
dieds.Interval = dieds.Interval * Timer2.Interval / 200
dieds.Enabled = True
TotalDied = TotalDied + 1
End Sub
Sub died2()
On Error Resume Next
di.Visible = False
dieds2.Interval = 20000
If TotalDied2 < 19 Then
dieds2.Interval = 1000 * TotalDied2 * 2 + 10000
Else
dieds2.Interval = 500000
End If
dieds2.Interval = dieds2.Interval * Timer2.Interval / 200
dieds2.Enabled = True
inf "敌方英雄已被杀害!"
diX = 基地2X
diHP = 0
diMP = 0
want2x = ""
TotalDied2 = TotalDied2 + 1
End Sub
Sub ab1()
'Exit Sub
For a = 1 To 20
If b11(a).e = False Then
b11(a).e = True
b11(a).hp = 500
b1(a).Left = -1000
b1(a).Visible = True
b11(a).X = 0
Exit For
Exit For
End If
Next
End Sub
Sub ab2()
For a = 1 To 20
If b22(a).e = False Then
b22(a).e = True
b22(a).hp = 500
b22(a).X = 基地2X
b2(a).Left = -1000
b2(a).Visible = True
Exit For
End If
Next
End Sub
Private Sub inf(str, Optional delay = 5)
If str = "该技能还未准备好！" Then Exit Sub
infs.Caption = str
infs.Visible = True
oinf = times + delay

End Sub


Private Sub abouts_Click()
inf "本游戏由YYZX-465-CH制作！" & vbCrLf & "游戏方法:QWER技能，B回城，" & vbCrLf & "左键普通攻击，右键走路"
End Sub

Private Sub bar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
tov = True
ElseIf Button = 2 Then

End If
End Sub


Private Sub bar_MouseMove(Button As Integer, Shift As Integer, X0 As Single, Y As Single)
If Rnd < 0.5 Then Exit Sub
bar.Print "小地图"
If tov = True Then
    xs = X0 / bar.Width * 基地2X - (Me.Width - ziji.Width) / 2
    基地1.Left = 基地1X - xs
    水晶1.Left = 水晶1X - xs
    塔1.Left = 塔1X - xs
    塔2.Left = 塔2X - xs
    塔3.Left = 塔3X - xs
    塔4.Left = 塔4X - xs
    水晶2.Left = 水晶2X - xs
    基地2.Left = 基地2X - xs
    di.Left = diX - xs
    ziji.Left = X - xs
    For a = 1 To 20
        If b11(a).e Then
            b1(a).Left = b11(a).X - xs
        End If
        If b22(a).e Then
            b2(a).Left = b22(a).X - xs
        End If
        If a <= 10 Then
            If daodan0(a).e Then
                daodan(a).Left = daodan0(a).X - xs
            End If
        End If
    Next
End If
End Sub

Private Sub bar_MouseUp(Button As Integer, Shift As Integer, X0 As Single, Y As Single)
If Button = 1 Then
tov = False
    x1 = X - (Me.Width - ziji.Width) / 2
    基地1.Left = 基地1X - x1
    水晶1.Left = 水晶1X - x1
    塔1.Left = 塔1X - x1
    塔2.Left = 塔2X - x1
    塔3.Left = 塔3X - x1
    塔4.Left = 塔4X - x1
    水晶2.Left = 水晶2X - x1
    基地2.Left = 基地2X - x1
    di.Left = diX - x1
    ziji.Left = X - x1
    For a = 1 To 20
        If b11(a).e Then
            b1(a).Left = b11(a).X - x1
        End If
        If b22(a).e Then
            b2(a).Left = b22(a).X - x1
        End If
        If a <= 10 Then
            If daodan0(a).e Then
                daodan(a).Left = daodan0(a).X - x1
            End If
        End If
    Next
End If
End Sub

Private Sub Command1_Click()
dd X, X + 3500, 0
End Sub

Private Sub Command10_Click()
diX = 0: TotalHP1 = 0
End Sub

Private Sub Command11_Click()
dd X, X + 3500, 1
End Sub

Private Sub Command12_Click()
dd X, X - 3500, 0
End Sub

Private Sub Command13_Click()
dd X, X - 3500, 1
End Sub

Private Sub Command14_Click()
c1 (Maxhp)
c2 (Maxmp)
End Sub


Private Sub Command16_Click()
wp2.Left = diX - X + (Me.Width - ziji.Width) / 2 - 5000
                'wp2.Left = (ziji.Left - di.Width) / 2
                wp2.Width = 10000
                wp2.Top = Y - wp2.Height - di.Height
                wp2.Visible = True
End Sub

Private Sub Command15_Click()
exp1 = exp1 + 50
End Sub

Private Sub Command17_Click()
updateview
Timer2_Timer
End Sub


Private Sub Command18_Click()
Call Text3_KeyPress(0)
End Sub

Private Sub Command2_Click()
ab1
End Sub

Private Sub Command3_Click()
ab2
End Sub

Private Sub Command4_Click()
inf times
End Sub

Private Sub Command5_Click()
oldxx = ""
Form_Activate

End Sub

Private Sub Command6_Click()
TotalHP2 = 0
End Sub






Private Sub Command7_Click()
Frame1.Visible = False
End Sub

Private Sub Command8_Click()
'Frame1.Visible = True
exp2 = exp2 + 100
End Sub

Private Sub Command9_Click()
X = 55000
End Sub


Private Sub dieds_Timer()
X = 0
c1 (Maxhp)
c2 (Maxmp)
ziji.Visible = True
dieds.Enabled = False
cana = 0
canq1 = 0
canw1 = 0
cane1 = 0
canr1 = 0
End Sub

Private Sub dieds2_Timer()
diX = 基地2X
diHP = diMaxHP
diMP = diMaxMP
inf "敌方已重生"
di.Visible = True
dieds2.Enabled = False
cana2 = 0
canq2 = 0
canw2 = 0
cane2 = 0
canr2 = 0
End Sub

Private Sub ends_Click()
End
End Sub
Private Sub toview(xs)

'Me.Cls
If Not isvb6 Then On Error GoTo err:
Dim 基地1Left, 水晶1Left, 塔1Left, 塔2Left, 塔3Left, 塔4Left, 水晶2Left, 基地2Left, diLeft, zijiLeft
If tov = False Then
    If 基地1X - xs + 基地1.Width >= -50 And 基地1X - xs <= Me.Width + 50 Then 基地1.Left = 基地1X - xs
    If 水晶1X - xs + 水晶1.Width >= -50 And 水晶1X - xs <= Me.Width + 50 Then 水晶1.Left = 水晶1X - xs
    If 塔1X - xs + 塔1.Width >= -50 And 塔1X - xs <= Me.Width + 50 Then 塔1.Left = 塔1X - xs
    If 塔2X - xs + 塔2.Width >= -50 And 塔2X - xs <= Me.Width + 50 Then 塔2.Left = 塔2X - xs
    If 塔3X - xs + 塔3.Width >= -50 And 塔3X - xs <= Me.Width + 50 Then 塔3.Left = 塔3X - xs
    If 塔4X - xs + 塔4.Width >= -50 And 塔4X - xs <= Me.Width + 50 Then 塔4.Left = 塔4X - xs
    If 水晶2X - xs + 水晶2.Width >= -50 And 水晶2X - xs <= Me.Width + 50 Then 水晶2.Left = 水晶2X - xs
    If 基地2X - xs + 基地2.Width >= -50 And 基地2X - xs <= Me.Width + 50 Then 基地2.Left = 基地2X - xs
    If diX - xs + di.Width >= -50 And diX - xs <= Me.Width + 50 Then di.Left = diX - xs
    ziji.Left = X - xs
End If
    基地1Left = 基地1X - xs
    水晶1Left = 水晶1X - xs
    塔1Left = 塔1X - xs
    塔2Left = 塔2X - xs
    塔3Left = 塔3X - xs
    塔4Left = 塔4X - xs
    水晶2Left = 水晶2X - xs
    基地2Left = 基地2X - xs
    diLeft = diX - xs
    zijiLeft = X - xs
    
For a = 1 To 20
Dim i As Boolean
i = False
If b11(a).e Then
    For bb = 1 To 20
    If b22(bb).e Then
        If b22(bb).X < b11(a).X + 1000 And b22(bb).X > b11(a).X Then
            i = True
            b22(bb).hp = b22(bb).hp - 1
                If b22(bb).hp <= 0 Then b22(bb).e = False: b2(bb).Visible = False: exp1 = exp1 + 0.5
        End If
    End If
    Next
    If 塔3X < b11(a).X + 1000 Then
        If b11(a).X < 塔3X Then
            If T3hp > 0 Then
                T3hp = T3hp - 1
                i = True
                If T3hp <= 0 Then 塔3.Picture = LoadResPicture(108, vbResBitmap): exp1 = exp1 + 5
            End If
        End If
    End If
    
    If 塔4X < b11(a).X + 1000 Then
        If b11(a).X < 塔4X Then
            If T4hp > 0 Then
                T4hp = T4hp - 1
                i = True
                If T4hp <= 0 Then 塔4.Picture = LoadResPicture(108, vbResBitmap): exp1 = exp1 + 5
            End If
        End If
    End If
    
    If 水晶2X < b11(a).X + 1000 Then
        If b11(a).X < 水晶2X Then
            i = True
            If Rnd < 0.5 Then
                If TotalHP2 > 0 Then
                    TotalHP2 = TotalHP2 - 1
                End If
            End If
        End If
    End If
    
    If diX < b11(a).X + 1000 Then
        If diX > b11(a).X Then
            i = True
            diHP = diHP - 1
            If diHP <= 0 Then died2: exp1 = exp1 + 8
        End If
    End If
    
    If Not i Then b11(a).X = b11(a).X + 30
    If tov = False Then
        If b11(a).X - xs + b1(a).Width >= -50 Then
            If b11(a).X - xs <= Me.Width + 50 Then
                b1(a).Left = b11(a).X - xs
            End If
        End If
    End If
End If

i = False

''''''''''''''''''''''''''''''''''''''''''''''

If b22(a).e Then
    For bb = 1 To 20
    If b11(bb).e Then
        If b11(bb).X > b22(a).X - 1000 Then
            If b11(bb).X < b22(a).X Then
                i = True
                b11(bb).hp = b11(bb).hp - 1
                    If b11(bb).hp <= 0 Then b11(bb).e = False: b1(bb).Visible = False: exp2 = exp2 + 0.5
            End If
        End If
    End If
    Next
    
    
    If 塔1X > b22(a).X - 1000 Then
    If b22(a).X > 塔1X Then
            If T1hp > 0 Then
                T1hp = T1hp - 1
                i = True
                If T1hp <= 0 Then 塔1.Picture = LoadResPicture(108, vbResBitmap): exp2 = exp2 + 5
            End If
        End If
    End If
    If 塔2X > b22(a).X - 1000 Then
        If b22(a).X > 塔2X Then
            If T2hp > 0 Then
                T2hp = T2hp - 1
                i = True
                If T2hp <= 0 Then 塔2.Picture = LoadResPicture(108, vbResBitmap): exp2 = exp2 + 5
            End If
        End If
    End If
    
    If 水晶1X > b22(a).X - 1000 Then
        If b22(a).X > 水晶1X Then
            If Rnd < 0.5 Then
                If TotalHP1 > 0 Then
                    TotalHP1 = TotalHP1 - 1
                End If
            End If
            i = True
        End If
    End If
    
    
    
    If X > b22(a).X - 1000 Then
        If X < b22(a).X Then
            i = True
            c1 (hp - 1)
            ox = "": Bs = 0: Bpic.Visible = False
        End If
    End If

    If Not i Then b22(a).X = b22(a).X - 30
    If tov = False Then
        If b22(a).X - xs + b2(a).Width >= -50 Then
            If b22(a).X - xs <= Me.Width + 50 Then
            b2(a).Left = b22(a).X - xs
            End If
        End If
    End If
End If
Next
i = False


If T1hp > 0 Then
    For a = 1 To 20
        If b22(a).e Then
            If 塔1X + 4000 > b22(a).X Then
                If 塔1X < b22(a).X Then
                    b22(a).hp = b22(a).hp - 1
                    l10 塔1.Left + 塔1.Width / 2, Y - 塔1.Height, b2(a).Left + b2(a).Width / 2, Y - b2(a).Height
                    If b22(a).hp <= 0 Then b22(a).e = False: b2(a).Visible = False: exp1 = exp1 + 1
                End If
            End If
            If 塔1X - 4000 > b22(a).X Then
                If 塔1X > b22(a).X Then
                    b22(a).hp = b22(a).hp - 1
                    l10 塔1.Left + 塔1.Width / 2, Y - 塔1.Height, b2(a).Left + b2(a).Width / 2, Y - b2(a).Height
                    If b22(a).hp <= 0 Then b22(a).e = False: b2(a).Visible = False: exp1 = exp1 + 1
                End If
            End If
        End If
    Next
    If Not i Then
        If 塔1X + 4000 > diX Then
            If 塔1X < diX Then
                diHP = diHP - 2
                i = True
                l10 塔1.Left + 塔1.Width / 2, Y - 塔1.Height, di.Left + di.Width / 2, Y - di.Height
                If diHP <= 0 Then died2: exp1 = exp1 + 5
            End If
        End If
        If 塔1X - 4000 < diX Then
            If 塔1X > diX Then
                diHP = diHP - 2
                i = True
                l10 塔1.Left + 塔1.Width / 2, Y - 塔1.Height, di.Left + di.Width / 2, Y - di.Height
                If diHP <= 0 Then died2: exp1 = exp1 + 5
            End If
        End If
    End If
End If

i = False
If T2hp > 0 Then
    For a = 1 To 20
        If b22(a).e Then
            If 塔2X + 4000 > b22(a).X Then
                If 塔2X < b22(a).X Then
                    b22(a).hp = b22(a).hp - 1
                    l20 塔2.Left + 塔2.Width / 2, Y - 塔2.Height, b2(a).Left + b2(a).Width / 2, Y - b2(a).Height
                    If b22(a).hp <= 0 Then b22(a).e = False: b2(a).Visible = False: exp1 = exp1 + 1
                End If
            End If
            If 塔2X - 4000 > b22(a).X Then
                If 塔2X > b22(a).X Then
                    b22(a).hp = b22(a).hp - 1
                    l20 塔2.Left + 塔2.Width / 2, Y - 塔2.Height, b2(a).Left + b2(a).Width / 2, Y - b2(a).Height
                    If b22(a).hp <= 0 Then b22(a).e = False: b2(a).Visible = False: exp1 = exp1 + 1
                End If
            End If
        End If
    Next
    If Not i Then
        If 塔2X + 4000 > diX Then
            If 塔2X < diX Then
                diHP = diHP - 2
                i = True
                l20 塔2.Left + 塔2.Width / 2, Y - 塔2.Height, di.Left + di.Width / 2, Y - di.Height
                If diHP <= 0 Then died2: exp1 = exp1 + 5
            End If
        End If
        If 塔2X - 4000 < diX Then
            If 塔2X > diX Then
                diHP = diHP - 2
                i = True
                l20 塔2.Left + 塔2.Width / 2, Y - 塔2.Height, di.Left + di.Width / 2, Y - di.Height
                If diHP <= 0 Then died2: exp1 = exp1 + 5
            End If
        End If
    End If
End If
'''''''''''''''''''''''''''''''''''''''''''''
i = False
If T3hp > 0 Then
    For a = 1 To 20
        If b11(a).e Then
            If 塔3X + 4000 > b11(a).X Then
                If 塔3X < b11(a).X Then
                    b11(a).hp = b11(a).hp - 1
                    i = True
                    l30 塔3.Left + 塔3.Width / 2, Y - 塔3.Height, b1(a).Left + b1(a).Width / 2, Y - b1(a).Height
                    If b11(a).hp <= 0 Then b11(a).e = False: b1(a).Visible = False: exp2 = exp2 + 1
                End If
            End If
            If 塔3X - 4000 < b11(a).X Then
                If 塔3X > b11(a).X Then
                    b11(a).hp = b11(a).hp - 1
                    i = True
                    l30 塔3.Left + 塔3.Width / 2, Y - 塔3.Height, b1(a).Left + b1(a).Width / 2, Y - b1(a).Height
                    If b11(a).hp <= 0 Then b11(a).e = False: b1(a).Visible = False: exp2 = exp2 + 1
                End If
            End If
        End If
    Next
    If Not i Then
        If 塔3X + 4000 > X Then
            If 塔3X < X Then
                c1 (hp - 2)
                l30 塔3.Left + 塔3.Width / 2, Y - 塔3.Height, ziji.Left + ziji.Width / 2, Y - ziji.Height
                ox = "": Bs = 0: Bpic.Visible = False
                i = True
                If hp <= 0 Then died: exp2 = exp2 + 5
            End If
        End If
        If 塔3X - 4000 < X Then
            If 塔3X > X Then
                c1 (hp - 2)
                l30 塔3.Left + 塔3.Width / 2, Y - 塔3.Height, ziji.Left + ziji.Width / 2, Y - ziji.Height
                ox = "": Bs = 0: Bpic.Visible = False
                i = True
                If hp <= 0 Then died: exp2 = exp2 + 5
            End If
        End If
    End If
End If

i = False
If T4hp > 0 Then
    For a = 1 To 20
        If b11(a).e Then
            If 塔4X + 4000 > b11(a).X Then
                If 塔4X < b11(a).X Then
                    b11(a).hp = b11(a).hp - 1
                    i = True
                    l40 塔4.Left + 塔4.Width / 2, Y - 塔4.Height, b1(a).Left + b1(a).Width / 2, Y - b1(a).Height
                    If b11(a).hp <= 0 Then b11(a).e = False: b1(a).Visible = False: exp2 = exp2 + 1
                End If
            End If
             If 塔4X - 4000 < b11(a).X Then
                 If 塔4X > b11(a).X Then
                    b11(a).hp = b11(a).hp - 1
                    i = True
                    l40 塔4.Left + 塔4.Width / 2, Y - 塔4.Height, b1(a).Left + b1(a).Width / 2, Y - b1(a).Height
                    If b11(a).hp <= 0 Then b11(a).e = False: b1(a).Visible = False: exp2 = exp2 + 1
                End If
            End If
        End If
    Next
    If Not i Then
        If 塔4X + 4000 > X Then
            If 塔4X < X Then
                c1 (hp - 2)
                ox = "": Bs = 0: Bpic.Visible = False
                l40 塔4.Left + 塔4.Width / 2, Y - 塔4.Height, ziji.Left + ziji.Width / 2, Y - ziji.Height
                i = True
                If hp <= 0 Then died: exp2 = exp2 + 5
            End If
        End If
        If 塔4X - 4000 < X Then
            If 塔4X > X Then
                c1 (hp - 2)
                ox = "": Bs = 0: Bpic.Visible = False
                l40 塔4.Left + 塔4.Width / 2, Y - 塔4.Height, ziji.Left + ziji.Width / 2, Y - ziji.Height
                i = True
                If hp <= 0 Then died: exp2 = exp2 + 5
            End If
        End If
    End If
End If

''''''''''''''''''''''''''''''导弹开始
For a = 1 To 10
    If daodan0(a).e Then
        If daodan0(a).f = 0 Then
        daodan0(a).X = daodan0(a).X - 200
            If daodan0(a).X < daodan0(a).want Then
                daodan0(a).e = False
                daodan(a).Visible = False
            End If
        ElseIf daodan0(a).f = 1 Then
        daodan0(a).X = daodan0(a).X + 200
            If daodan0(a).X > daodan0(a).want Then
                daodan0(a).e = False
                daodan(a).Visible = False
            End If
        End If
        
        
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If Left(daodan0(a).ty, 1) = "." Or Left(daodan0(a).ty, 1) = "0" Then
                For B = 1 To 20
                    If b22(B).e Then
                        If b22(B).X < daodan0(a).X + daodan(a).Width - ziji.Left Then
                            If b22(B).X + b2(B).Width + ziji.Left > daodan0(a).X Then
                                daodan0(a).e = False
                                daodan(a).Visible = False
                                If daodan0(a).ty = 0.2 Then
                                    If b22(B).hp < 100 Then
                                        b22(B).hp = b22(B).hp - 550
                                        daodan0(a).e = True
                                        daodan(a).Visible = True
                                    Else
                                        b22(B).hp = b22(B).hp - 550
                                    End If
                                    
                                ElseIf daodan0(a).ty = 0.1 Then
                                    b22(B).hp = b22(B).hp - 250
                                Else
                                    b22(B).hp = b22(B).hp - 100
                                End If
                                xx "小兵"
                                bingid = B
                                
                                If b22(B).hp <= 0 Then
                                    b2(B).Visible = False
                                    b22(B).e = False
                                    exp1 = exp1 + 5
                                End If
                            End If
                        End If
                    End If
                Next
                If 塔3X < daodan0(a).X + daodan(a).Width - ziji.Left Then
                    If 塔3X + 塔3.Width + ziji.Left > daodan0(a).X Then
                        If T3hp > 0 Then
                            If daodan0(a).ty = 0.2 Then
                                T3hp = T3hp - 350
                            ElseIf daodan0(a).ty = 0.1 Then
                                T3hp = T3hp - 120
                            Else
                                T3hp = T3hp - 50
                            End If
                            daodan0(a).e = False
                            daodan(a).Visible = False
                            xx "塔3"
                            If T3hp <= 0 Then
                                塔3.Picture = LoadResPicture(108, vbResBitmap)
                                exp1 = exp1 + 10
                            End If
                        End If
                    End If
                End If
                If 塔4X < daodan0(a).X + daodan(a).Width - ziji.Left Then
                    If 塔4X + 塔4.Width + ziji.Left > daodan0(a).X Then
                        If T4hp > 0 Then
                            If daodan0(a).ty = 0.2 Then
                                T4hp = T4hp - 350
                            ElseIf daodan0(a).ty = 0.1 Then
                                T4hp = T4hp - 120
                            Else
                                T4hp = T4hp - 50
                            End If
                            daodan0(a).e = False
                            daodan(a).Visible = False
                            xx "塔4"
                            If T4hp <= 0 Then
                                塔4.Picture = LoadResPicture(108, vbResBitmap)
                                exp1 = exp1 + 10
                            End If
                        End If
                    End If
                End If
                If diHP > 0 Then
                    If diX < daodan0(a).X + daodan(a).Width - ziji.Left Then
                        If diX + di.Width + ziji.Left > daodan0(a).X Then
                            If daodan0(a).ty = 0.2 Then
                                diHP = diHP - 350
                                downv = 35 + lev2
                            ElseIf daodan0(a).ty = 0.1 Then
                                diHP = diHP - 120
                            Else
                                diHP = diHP - 50
                            End If
                            xx "敌人"
                                daodan0(a).e = False
                                daodan(a).Visible = False
                            If diHP <= 0 Then died2: exp1 = exp1 + 15
                        End If
                    End If
                End If
                If 水晶2X < daodan0(a).X + daodan(a).Width - ziji.Left Then
                    If 水晶2X + 水晶2.Width + ziji.Left > daodan0(a).X Then
                        If daodan0(a).ty = 0.2 Then
                            TotalHP2 = TotalHP2 - 300
                        ElseIf daodan0(a).ty = 0.1 Then
                            TotalHP2 = TotalHP2 - 150
                        Else
                            TotalHP2 = TotalHP2 - 50
                        End If
                        xx "水晶2"
                        daodan0(a).e = False
                        daodan(a).Visible = False
                    End If
                End If
                ''''''''''''''''''''''''''''''''''''''
            ElseIf Left(daodan0(a).ty, 1) = "1" Then
                For B = 1 To 20
                    If b11(B).e Then
                        If b11(B).X < daodan0(a).X + daodan(a).Width - ziji.Left Then
                            If b11(B).X + b1(B).Width + ziji.Left > daodan0(a).X Then
                                daodan0(a).e = False
                                daodan(a).Visible = False
                                If daodan0(a).ty = 1.2 Then
                                    b11(B).hp = b11(B).hp - 450
                                ElseIf daodan0(a).ty = 1.1 Then
                                    b11(B).hp = b11(B).hp - 250
                                Else
                                    b11(B).hp = b11(B).hp - 100
                                End If
                                If b11(B).hp <= 0 Then
                                    b1(B).Visible = False
                                    b11(B).e = False
                                    exp2 = exp2 + 5
                                End If
                            End If
                        End If
                    End If
                Next
                If 塔1X < daodan0(a).X + daodan(a).Width - ziji.Left And 塔1X + 塔1.Width + ziji.Left > daodan0(a).X Then
                    If T1hp > 0 Then
                        If daodan0(a).ty = 1.2 Then
                            T1hp = T1hp - 350
                        ElseIf daodan0(a).ty = 1.1 Then
                            T1hp = T1hp - 120
                        Else
                            T1hp = T1hp - 50
                        End If
                        daodan0(a).e = False
                        daodan(a).Visible = False
                        If T1hp <= 0 Then
                            塔1.Picture = LoadResPicture(108, vbResBitmap)
                            exp2 = exp2 + 10
                        End If
                    End If
                End If
                If 塔2X < daodan0(a).X + daodan(a).Width - ziji.Left Then
                    If 塔2X + 塔2.Width + ziji.Left > daodan0(a).X Then
                        If T2hp > 0 Then
                            If daodan0(a).ty = 1.2 Then
                                T2hp = T2hp - 350
                            ElseIf daodan0(a).ty = 1.1 Then
                                T2hp = T2hp - 120
                            Else
                                T1hp = T2hp - 50
                            End If
                            daodan0(a).e = False
                            daodan(a).Visible = False
                            If T2hp <= 0 Then
                                塔2.Picture = LoadResPicture(108, vbResBitmap)
                                exp2 = exp2 + 10
                            End If
                        End If
                    End If
                End If
                If X < daodan0(a).X + daodan(a).Width - ziji.Left Then
                    If X + ziji.Width + ziji.Left > daodan0(a).X Then
                        If daodan0(a).ty = 1.2 Then
                            c1 (hp - 350)
                            inf "禁锢！", 2
                            downvme = 35 + lev1
                        ElseIf daodan0(a).ty = 1.1 Then
                            c1 (hp - 120)
                        Else
                            c1 (hp - 50)
                        End If
                        ox = "": Bs = 0: Bpic.Visible = False
                        daodan0(a).e = False
                        daodan(a).Visible = False
                        If hp <= 0 Then died: exp2 = exp2 + 15
                    End If
                End If
                If 水晶1X < daodan0(a).X + daodan(a).Width - ziji.Left Then
                    If 水晶1X + 水晶1.Width + ziji.Left > daodan0(a).X Then
                        If daodan0(a).ty = 1.2 Then
                            TotalHP1 = TotalHP1 - 300
                        ElseIf daodan0(a).ty = 1.1 Then
                            TotalHP1 = TotalHP1 - 150
                        Else
                            TotalHP1 = TotalHP1 - 50
                        End If
                        daodan0(a).e = False
                        daodan(a).Visible = False
                    End If
                End If
            End If
        
        If tov = False Then
            If daodan0(a).X - X + daodan(a).Width >= -50 Then
                If daodan0(a).X - X <= Me.Width + 50 Then
                    daodan(a).Left = daodan0(a).X - X
                End If
            End If
        End If
    End If
Next
''''''''''''''''''''导弹结束


If wantx <> "" Then
If wantx < X Then
X = X - 35 - lev1 + downvme
fang = 0
If wantx >= X Then wantx = ""
ElseIf wantx > X Then
X = X + 35 + lev1 - downvme
fang = 1
If wantx <= X Then wantx = ""
End If
End If

If want2x <> "" Then
    If want2x < diX Then
        diX = diX - 35 - lev2 + downv
            If want2x >= diX Then want2x = ""
    ElseIf want2x > diX Then
        diX = diX + 35 + lev2 - downv
            If want2x <= X Then want2x = ""
    End If
End If
If downv <> 0 Then
If Rnd < 0.1 Then
downv = downv - 1
End If
End If
If downvme <> 0 Then
If Rnd < 0.1 Then
downvme = downvme - 1
End If
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Text1 = diHP
If diHP > 0 Then
Text1 = "ing"
    Dim fangan As Long, pchp, userhp
    userhp = hp / Maxhp
    pchp = diHP / diMaxHP
    If pcback = "" Then fangan = 1
    If Int(diX) < Int(b22(10).X) Then
        fangan = 0
    End If
    For a = 1 To 20
        If b11(a).e Then
            If b11(a).X + b1(a).Width > diX - 3500 Then
                If b11(a).X < diX Then
                    If cana2 <= times Then
                        fangan = 2
                    End If
                    If Rnd < 0.2 Then
                        fangan = 5
                    End If
                End If
            End If
            If diX + di.Width + 3500 > b11(a).X Then
                If b11(a).X > diX Then
                    If cana2 <= times Then
                        fangan = 3
                    End If
                    If Rnd < 0.2 Then
                        fangan = 6
                    End If
                End If
             End If
        End If
        Next
        If cana2 <= times Then
            If 水晶1X + 水晶1.Width > diX - 3500 Then
                If 水晶1X < diX Then
                    fangan = 2
                End If
            End If
        End If
        If cana2 <= times Then
            If diX + di.Width + 3500 > 水晶1X Then
                If 水晶1X > diX Then
                    fangan = 3
                End If
            End If
        End If
        If T1hp > 0 Then
            If cana2 <= times Then
                If 塔1X + 塔1.Width > diX - 3500 Then
                    If 塔1X < diX Then
                        fangan = 2
                    End If
                End If
            End If
            If cana2 <= times Then
                If diX + di.Width + 3500 > 塔1X Then
                    If 塔1X > diX Then
                        fangan = 3
                    End If
                End If
            End If
        End If
        If T2hp > 0 Then
            If cana2 <= times Then
                If 塔2X + 塔2.Width > diX - 3500 Then
                    If 塔2X < diX Then
                        fangan = 2
                    End If
                End If
            End If
            If cana2 <= times Then
                If diX + di.Width + 3500 > 塔2X Then
                    If 塔2X > diX Then
                        fangan = 3
                    End If
                End If
            End If
        End If
        If X + ziji.Width > diX - 3500 Then
            If X < diX Then
                fangan = 2
                If canq2 <= times Then
                    fangan = 5
                End If
                If canr2 <= times Then
                    fangan = 9
                End If
            End If
        End If
        If diX + di.Width + 3500 > X Then
            If X > diX Then
                fangan = 3
                If canq2 <= times Then
                    fangan = 6
                End If
                If canr2 <= times Then
                    fangan = 10
                End If
            End If
        End If
        
    
    If diMaxHP - 300 > diHP And diMaxMP - 300 > diMP Then
        If cane2 <= times Then
            If lev2 >= 6 Then
                If diMP - 1 >= 0 Then
                    fangan = 8
                End If
            End If
        End If
    End If
    If pchp < 0.2 Then
        If pchp < userhp Then
            If diX < 基地2X Then
                pcback = 1
                If Rnd < 0.9 Then fangan = 4
            End If
        End If
    End If
    If pchp < 0.05 Then
        If diX < 基地2X Then
            pcback = 1
            If Rnd < 0.9 Then fangan = 4
        End If
    End If
    If diMP < 0 Then
        If diX < 基地2X Then
            pcback = 1
            If Rnd < 0.05 Then fangan = 4
        End If
    End If
    If hp <= 0 Then
        If diHP < 390 Then
            pcback = 1
            If Rnd < 0.5 Then fangan = 4
        End If
    End If
    
    If fangan = 4 Then
    If cane2 <= times Then
    If lev2 >= 6 Then
        If Rnd < 0.05 Then fangan = 8
    End If
    End If
    If canw2 <= times Then
    If Rnd < 0.05 Then fangan = 7
    End If
    End If
    If Text2 <> "" Then fangan = Text2
    Text1 = fangan
    'If Text1 <> "" Then fangan = Text1
    If fangan = 2 Then
    If Rnd < 0.005 Then fangan = 9
    End If
    If fangan = 3 Then
    If Rnd < 0.005 And pcback = "" Then fangan = 1
    End If
    If Rnd < 0.005 Then fangan = 4
    Select Case fangan
    Case "1" '前进
        ditox diX - 1000
    Case "2" '左射击
zsj:
        If cana2 <= times Then
            cana2 = times + 1
            dd diX, diX - 3500, 1
        End If
    Case "3" '右射击
ysj:
        If cana2 < times Then
            cana2 = cana2 + 1
            dd diX, diX + 3500, 1
        End If
    Case "4" '逃跑
        ditox diX + 10
    Case "5" 'q 50 left
    If canq2 <= times Then
        If lev2 >= 2 Then
            If diMP - 50 >= 0 Then
                canq2 = times + 3 - lev2 / 9
                diMP = diMP - 50
                dd diX, diX - 10000, 1.1
            Else
            fangan = 2
            GoTo zsj:
            End If
        End If
    End If
    Case "6" 'q 50 right
    If canq2 <= times Then
        If lev2 >= 2 Then
            If diMP - 50 >= 0 Then
                canq2 = times + 3 - lev2 / 9
                diMP = diMP - 50
                dd diX, diX + 10000, 1.1
            Else
            fangan = 3
            GoTo ysj:
            End If
        End If
    End If
    Case "7" 'w 100
    If canw2 <= times Then
        If lev2 >= 4 Then
            If diMP - 100 >= 0 Then
                canw2 = times + 10 - lev2 / 4
                diMP = diMP - 100
        'canw1 = times + 10 - lev1 / 4
                wp2.Left = diX - X + (Me.Width - ziji.Width) / 2 - 5000
                wp2.Width = 10000
                wp2.Top = Y - wp2.Height - di.Height
                wp2.Visible = True
    'End If
                'diMP = diMP - 60
                'dd diX, diX + 10000, 1.1
                'dd diX, diX - 10000, 1.1
                'dd diX, diX + 20000, 1.1
                'dd diX, diX - 20000, 1.1
            End If
        End If
    End If
    Case "8" 'e 250
    If cane2 <= times Then
        If lev2 >= 6 Then
            If diMP - 1 >= 0 Then
                cane2 = times + 20 - lev2 / 2
                diMP = diMP + 300
                diHP = diHP + 300
            End If
        End If
    End If
    Case "9" 'r 400 left
        If canr2 <= times Then
            If lev2 >= 8 Then
                canr2 = times + 30 - lev2
                diMP = diMP - 400
                dd diX, diX - 30000, 1.2
                dd diX, diX - 30000, 1.2
                dd diX, diX - 10000, 1
                dd diX, diX - 7000, 1
                dd diX, diX - 3000, 1
                Else
                fangan = 2
                GoTo zsj:
            End If
        End If
    Case "10" 'r 400 right
        If canr2 <= times Then
            If lev2 >= 8 Then
                canr2 = times + 30 - lev2
                diMP = diMP - 400
                dd diX, diX + 30000, 1.2
                dd diX, diX + 30000, 1.2
                dd diX, diX + 10000, 1
                dd diX, diX + 7000, 1
                dd diX, diX + 3000, 1
                Else
                fangan = 3
                GoTo ysj:
            End If
        End If
    End Select
End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''


If wp.Visible = True Then
wp.Top = wp.Top + 60
    If wp.Top >= Y - wp.Height Then
    wp.Visible = False
        For a = 1 To 20
            If X + 5000 > b22(a).X Then
                If X < b22(a).X Then
                    b22(a).hp = b22(a).hp - 200
                    If b22(a).hp < 0 Then b22(a).e = False: b2(a).Visible = False
                End If
            End If
            If X - 5000 > b22(a).X Then
                If X > b22(a).X Then
                    b22(a).hp = b22(a).hp - 200
                    If b22(a).hp < 0 Then b22(a).e = False: b2(a).Visible = False
                End If
            End If
        Next
        If X + 5000 > diX Then
            If X < diX Then
                diHP = diHP - 400
                If diHP <= 0 Then died2: exp1 = exp1 + 5
            End If
        End If
        If X - 5000 > diX Then
            If X > diX Then
                diHP = diHP - 400
                If diHP <= 0 Then died2: exp1 = exp1 + 5
            End If
        End If
    End If
End If

If wp2.Visible = True Then
'wp2.Left =
wp2.Left = diX - X + (Me.Width - ziji.Width) / 2 - 5000
wp2.Top = wp2.Top + 60
    If wp2.Top >= Y - wp2.Height Then
    wp2.Visible = False
        For a = 1 To 20
            If X + 5000 > b11(a).X Then
                If X < b11(a).X Then
                    b11(a).hp = b11(a).hp - 200
                    If b11(a).hp < 0 Then b11(a).e = False: b1(a).Visible = False
                End If
            End If
            If diX - 5000 > b11(a).X Then
                If diX > b11(a).X Then
                    b11(a).hp = b11(a).hp - 200
                    If b11(a).hp < 0 Then b11(a).e = False: b1(a).Visible = False
                End If
            End If
        Next
        If diX + 5000 > X Then
            If X < diX Then
                'diHP = diHP - 400
                c1 (hp - 400)
                'If diHP <= 0 Then died2: exp1 = exp1 + 5
            End If
        End If
        If diX - 5000 > X Then
            If X > diX Then
                c1 (hp - 400)
                'diHP = diHP - 400
                'If HP <= 0 Then died2: exp1 = exp1 + 5
            End If
        End If
    End If
End If





''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If rehp > 0 Then
rehp = rehp - 1
c1 (hp + 3)
End If
If TotalHP1 <= 0 Then
Timer1.Enabled = False
Timer2.Enabled = False
'Command5.Visible = True
over.Enabled = True
End If
If TotalHP2 <= 0 Then
Timer1.Enabled = False
Timer2.Enabled = False
'Command5.Visible = True
over.Enabled = True
End If
'GdiTransparentBlt Me.hDC, -x / 15, y - 塔1.Height + 90, 塔1.ScaleWidth, 塔1.ScaleHeight, 塔1.hDC, 0, 0, 塔1.ScaleWidth, 塔1.ScaleHeight, RGB(255, 255, 255)
'Me.Refresh
Exit Sub
err:
inf "警告：赛先生联盟动作处理发生严重错误！"
End Sub

Private Sub Form_Activate()
If Command5.Visible = False Then Exit Sub
Form2.Show
'Me.Show
'On Error GoTo errload:
exp1 = 0
exp2 = 0
Me.AutoRedraw = False
If Not isvb6 Then
Frame1.Visible = False
zuobi.Visible = False
End If
If Dir("C:\zuobi.txt", vbNormal) <> "" Then
zuobi.Visible = True
End If
'infs.Caption = "游戏马上开始！游戏部署中..."
infs.Top = (Y - infs.Height) / 2 - 800
infs.Left = (Me.Width - infs.Width) / 4
inf "游戏马上开始！游戏部署中..."
ziji.Width = 1260
ziji.Height = 1665
Set jian1 = LoadResPicture(103, vbResBitmap) 'LoadPicture(AppPath & "\pic\j1.bmp")
Set jian2 = LoadResPicture(104, vbResBitmap) 'LoadPicture(AppPath & "\pic\j2.bmp")
塔1.Picture = LoadResPicture(107, vbResBitmap) 'LoadPicture(AppPath & "\pic\t2.bmp")
塔2.Picture = LoadResPicture(107, vbResBitmap) 'LoadPicture(AppPath & "\pic\t2.bmp")
塔3.Picture = LoadResPicture(107, vbResBitmap) 'LoadPicture(AppPath & "\pic\t2.bmp")
塔4.Picture = LoadResPicture(107, vbResBitmap) 'LoadPicture(AppPath & "\pic\t2.bmp")
水晶1.Picture = LoadResPicture(106, vbResBitmap) 'LoadPicture(AppPath & "\pic\t1.bmp")
水晶2.Picture = LoadResPicture(106, vbResBitmap) 'LoadPicture(AppPath & "\pic\t1.bmp")
di.Picture = LoadResPicture(105, vbResBitmap) 'LoadPicture(AppPath & "\pic\di.bmp")
For a = 1 To 20
    If b11(a).e Then
        b11(a).e = False
        b1(a).Visible = False
    End If
    If b22(a).e Then
        b22(a).e = False
        b2(a).Visible = False
    End If
    If a <= 10 Then
        If daodan0(a).e Then
            daodan0(a).e = False
            daodan(a).Visible = False
        End If
    End If
Next
水晶1.Width = 2265
水晶2.Width = 2265
ziji.BorderStyle = 0
di.BorderStyle = 0
基地1.BorderStyle = 0
水晶1.BorderStyle = 0
塔1.BorderStyle = 0
塔2.BorderStyle = 0
塔3.BorderStyle = 0
塔4.BorderStyle = 0
水晶2.BorderStyle = 0
基地2.BorderStyle = 0
cana = 0
cana2 = 0
fang = 1
水晶1.BackColor = vbWhite
水晶2.BackColor = vbWhite
Command5.Visible = False
X = 0
updateview
wantx = ""
reb = 0
times = 0
Timer1.Enabled = True
Timer2.Enabled = True
fpss = 0
fpss2 = Int(GetTickCount / 1000)
ziji.BackColor = vbWhite
基地1X = 0
水晶1X = 5000
塔1X = 10000
塔2X = 20000
塔3X = 40000
塔4X = 50000
水晶2X = 60000
基地2X = 70000
Maxhp = 1000
Maxmp = 1000
hp = 1000
MP = 1000
diHP = 1000
diMP = 1000
diMaxHP = 1000
diMaxMP = 1000
T1hp = 2000
T2hp = 2000
T3hp = 2000
T4hp = 2000
TotalHP1 = 5000
TotalHP2 = 5000
塔1.BackColor = vbWhite
塔2.BackColor = vbWhite
塔3.BackColor = vbWhite
塔4.BackColor = vbWhite
塔1.Width = 1330
塔2.Width = 1330
塔3.Width = 1330
塔4.Width = 1330

HPT.Height = 500
HPT.Width = Me.Width / 2
MPT.Height = 500
MPT.Width = Me.Width / 2
HPT.Left = Me.Width / 2 - HPT.Width / 2
HPT.Top = Me.Height - MPT.Height - HPT.Height
MPT.Left = HPT.Left
MPT.Top = Me.Height - MPT.Height
HPT.BackColor = vbRed
MPT.BackColor = vbBlue
bar.BackColor = vbYellow
bar.Left = 0
bar.Height = MPT.Height
bar.Width = Me.Width
bar.Top = HPT.Top - bar.Height
jipicwidth = 1000
jipicheight = 1000
qpic.Width = jipicwidth
wpic.Width = jipicwidth
epic.Width = jipicwidth
rpic.Width = jipicwidth
qpic.Height = jipicheight
wpic.Height = jipicheight
epic.Height = jipicheight
rpic.Height = jipicheight
qpic.Top = Me.Height - 2 * bar.Height - HPT.Height - qpic.Height - Label1.Height
wpic.Top = qpic.Top
epic.Top = qpic.Top
rpic.Top = qpic.Top
qpic.Left = Me.Width / 2 - 2 * wpic.Width
wpic.Left = Me.Width / 2 - wpic.Width
epic.Left = Me.Width / 2
rpic.Left = Me.Width / 2 + wpic.Width



Label1.Width = HPT.Width
Label1.Height = HPT.Height
Label1.Font.Size = 10
Label1.Top = bar.Top - Label1.Height
Label1.Left = 0 'HPT.Left


基地1.Left = 基地1X
水晶1.Left = 水晶1X
塔1.Left = 塔1X
塔2.Left = 塔2X
塔3.Left = 塔3X
塔4.Left = 塔4X
水晶2.Left = 水晶2X
基地2.Left = 基地2X


基地1.Height = 基地高
水晶1.Height = 水晶高
塔1.Height = 塔高
塔2.Height = 塔高
塔3.Height = 塔高
塔4.Height = 塔高
水晶2.Height = 水晶高
基地2.Height = 基地高
基地1.Top = Y - 基地高
水晶1.Top = Y - 水晶高
塔1.Top = Y - 塔高
塔2.Top = Y - 塔高
塔3.Top = Y - 塔高
塔4.Top = Y - 塔高
水晶2.Top = Y - 水晶高
基地2.Top = Y - 基地高
ziji.Picture = LoadResPicture(105, vbResBitmap) 'LoadPicture(AppPath & "\pic\r.bmp")
ziji.Top = Y - ziji.Height
di.Width = ziji.Width
di.Height = ziji.Height
di.BackColor = vbWhite
di.Top = ziji.Top
diX = 基地2X
di.Left = diX
'塔1.BackColor = transcolor
'Me.ScaleMode = 3
'塔2.Picture = Tpic
'塔3.Picture = Tpic
'塔4.Picture = Tpic
'GdiTransparentBlt Me.hDC, 0, 0, 1000, 1000, 塔1.hDC, 0, 0, 1000, 1000, RGB(0, 0, 0)
'AlphaBlend Me.hDC, 0, 0, 1000, 1000, 塔1.hDC, 0, 0, 1000, 1000, RGB(0, 0, 0)
'Image1.Container = Tpic
b22(10).X = 基地2X
Bpic.Left = 0
Bpic.Width = 280
Bpic.Height = 0
Bpic.Top = Y - Bpic.Height
Bpic.BackColor = vbWhite
Command5.Left = (Me.Width - Command5.Width) / 2
Command5.Top = Y / 2
canq1 = 0
canw1 = 0
cane1 = 0
canr1 = 0
canq2 = 0
canw2 = 0
cane2 = 0
canr2 = 0
塔2.Left = 塔2X - X + (Me.Width - ziji.Width) / 2
TotalDied = 0
ends.Left = Me.Width - ends.Width
ends.Top = 0
abouts.Left = ends.Left - abouts.Width
abouts.Top = 0
zuobi.Top = 0
zuobi.Left = abouts.Left - zuobi.Width
fpscpu = False
Text3.Text = 64
qpic.ToolTipText = "Q键（普通攻击）2级——向移动方向发射一只箭，造成150点伤害！"
wpic.ToolTipText = "W键（普通攻击）4级——造成大范围物理伤害，仅对小兵及敌方英雄有效！"
epic.ToolTipText = "E键（辅助技能）6级——为自己补充300点血！"
rpic.ToolTipText = "E键（魔法攻击）8级——向移动方向发射一只魔法水晶箭，当击中的小兵血量低于200点时，将杀死小兵继续向前飞！对英雄造成减速效果！"
TotalDied2 = 0
didi.Top = qpic.Top - didi.Height
didi.Left = qpic.Left
didi2.Top = didi.Top - didi2.Height
didi2.Left = didi.Left
fpscpu = False
toview -(Me.Width - ziji.Width) / 2

Dim rtn As Long
Me.BackColor = vbWhite
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, vbWhite, 0, LWA_COLORKEY
inf "10秒后动身！"

Exit Sub
errload:
MsgBox "赛先生联盟，启动过程中发生严重错误，请联系CH。错误代号：" & err.Number
End
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 37 Then key = 1
'If KeyCode = 39 Then key = 2
If hp <= 1 Then inf "你现在还不能发技能！", 3: Exit Sub

'81 87 69 82
'B 66
If GetAsyncKeyState(vbKeyQ) <> 0 Then
    If lev1 < 2 Then inf "该技能需达到2级才学会！", 3: Exit Sub
    If canq1 > times Then inf "该技能还未准备好！", 3: Exit Sub
    If c2(MP - 50) Then
    canq1 = times + 3 - lev1 / 11
        'If fang = 0 Then
        '    dd X, X - 10000, 0.1
        'Else
        '    dd X, X + 10000, 0.1
        'End If
        GetCursorPos p
        'MsgBox p.X
        If p.X * 15 < Screen.Width / 2 Then
            dd X, X - 10000, 0.1
        Else
            dd X, X + 10000, 0.1
        End If
        
        
    'For a = 1 To 20
    '    If (x + 5000 > b22(a).x And x < b22(a).x) Or (x - 5000 > b22(a).x And x > b22(a).x) Then
    '    b22(a).hp = b22(a).hp - 30
    '    If b22(a).hp < 0 Then b22(a).e = False: b2(a).Visible = False
    '    End If
    'Next
    End If
End If

If GetAsyncKeyState(vbKeyW) <> 0 Then
    If lev1 < 4 Then inf "该技能需达到4级才学会！", 3: Exit Sub
    If canw1 > times Then inf "该技能还未准备好！", 3: Exit Sub
        If c2(MP - 100) Then
        canw1 = times + 10 - lev1 / 4
        wp.Left = (ziji.Left - ziji.Width) / 2
        wp.Width = 10000
        wp.Top = Y - wp.Height - ziji.Height
        wp.Visible = True
    End If
End If

If GetAsyncKeyState(vbKeyE) <> 0 Then
    If lev1 < 6 Then inf "该技能需达到6级才学会！", 3: Exit Sub
    If cane1 > times Then inf "该技能还未准备好！", 3: Exit Sub
    If c2(MP - 1) Then
    cane1 = times + 20 - lev1 / 2
        rehp = rehp + 100
        MP = MP + 300
    End If
End If

If GetAsyncKeyState(vbKeyR) <> 0 Then
    If lev1 < 8 Then inf "该技能需达到8级才学会！", 3: Exit Sub
    If canr1 > times Then inf "该技能还未准备好！", 3: Exit Sub
    If c2(MP - 400) Then
    canr1 = times + 30 - lev1
        If fang = 0 Then
            dd X, X - 30000, 0.2
            dd X, X - 3500, 0
            dd X, X - 7000, 0
            dd X, X - 10000, 0
        Else
            dd X, X + 30000, 0.2
            dd X, X + 3500, 0
            dd X, X + 7000, 0
            dd X, X + 10000, 0
        End If
    End If

End If



If KeyCode = 66 Then wantx = ""
If KeyCode = 66 And X <> 0 Then
Bpic.Visible = True
If ox = "" Then ox = X
    If X <> ox Then
        Bpic.Height = 0: Bpic.Top = Y: ox = "": Bs = 0: Bpic.Visible = False
    Else
        Bpic.Height = Bpic.Height + fps / 0.64
        Bpic.Top = Y - Bpic.Height
        If Bpic.Height >= Y Then
                Bpic.Height = 0
                Bpic.Top = Y
                ox = ""
                Bs = 0
                Bpic.Visible = False
                X = 0
                updateview
        End If
    End If
        'If x <> ox Then
        '    Bpic.Height = 0: Bpic.Top = y: ox = "": Bs = 0: Bpic.Visible = False
        'Else
        '    Bpic.Top = Bpic.Top - fps / 64
        '    Bpic.Height = y - Bpic.Top
        '    'Bs = Int(Bs) + fps / 64
        '    'Bpic.Top = y - Bpic.Height
        '    'Bpic.Height = y * Bs / 100
        '    If Bpic.Height >= Me.Height Then
        '        Bpic.Height = 0
        '        Bpic.Top = y
        '        ox = ""
        '        Bs = 0
        '        Bpic.Visible = False
        '        x = 0
        '        updateview
        '    End If
        'End If
End If
End Sub

Private Sub Form_Load()
On Error GoTo err2:
Dim xb As Picture, xb2 As Picture
Set xb = LoadResPicture(109, vbResBitmap) 'LoadPicture(AppPath & "\pic\xb.bmp")
Set xb2 = LoadResPicture(110, vbResBitmap) 'LoadPicture(AppPath & "\pic\xb2.bmp")
For a = 1 To 20
Set b1(a) = Controls.Add("vb.picturebox", "b" & a)
b11(a).e = False
b1(a).Visible = False
b1(a).Width = 500
b1(a).Height = 500
b1(a).BackColor = vbWhite
b1(a).Top = Y - b1(a).Height
b1(a).BorderStyle = 0
b1(a).Picture = xb
Next
For a = 1 To 20
Set b2(a) = Controls.Add("vb.picturebox", "b2" & a)
b22(a).e = False
b2(a).Visible = False
b2(a).Width = 500
b2(a).Height = 500
b2(a).BackColor = vbWhite
b2(a).Top = Y - b2(a).Height
b2(a).BorderStyle = 0
b2(a).Picture = xb2
Next
For a = 1 To 10
Set daodan(a) = Controls.Add("vb.picturebox", "daodan" & a)
daodan(a).Visible = False
daodan(a).Left = -500
daodan(a).Width = 900
daodan(a).Height = 200
daodan(a).Top = Y - daodan(a).Height - ziji.Width / 4
daodan(a).Picture = jian1
daodan0(a).e = False
daodan(a).BorderStyle = 0
Next
Exit Sub
err2:
MsgBox "游戏加载时发生致命的错误，游戏所需要的文件不能被找到，请重新弄一下！——CH", vbCritical
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X0 As Single, Y As Single)
If Button = 1 Then
    If dieds.Enabled = False Then
        If cana <= times Then
            cana = times + 1
            If X0 > ziji.Left Then
            dd X, X + 3500, 0
            Else
            dd X, X - 3500, 0
            End If
        End If
    End If
End If
If Button = 2 Then
If dieds.Enabled = False Then
wantx = X0 + X - ziji.Left
End If
End If
Bpic.Height = 0: Bpic.Top = Y: ox = "": Bs = 0: Bpic.Visible = False
End Sub





Private Sub fpsshow_Click()
Timer1.Enabled = False
fpscpu = True
For a = 0 To 1
a = 0
Timer1_Timer
If fpscpu = False Then a = 1
Next
End Sub

Private Sub keys_Timer()
If GetAsyncKeyState(1) Then
GetCursorPos p
Form_MouseDown 1, 0, p.X * 15, 0
End If

If GetAsyncKeyState(2) Then
GetCursorPos p
Form_MouseDown 2, 0, p.X * 15, 0
End If
If GetAsyncKeyState(vbKeyQ) Then
Form_KeyDown 81, 0
End If
If GetAsyncKeyState(vbKeyW) Then
Form_KeyDown 87, 0
End If
If GetAsyncKeyState(vbKeyE) Then
Form_KeyDown 69, 0
End If
If GetAsyncKeyState(vbKeyR) Then
Form_KeyDown 82, 0
End If
'Me.Cls
'Me.Print GetAsyncKeyState(vbKeyQ)
End Sub

Private Sub over_Timer()
Text3.Text = 64
fpscpu = False

If oldxx = "" Then oldxx = X ': x = x - ziji.Left
ziji.Top = Me.Height + 5000
di.Top = Me.Height + 5000
If TotalHP1 <= 0 Then
    If X + Me.Width / 2 < 水晶1X + 水晶1.Width + 200 And X + Me.Width / 2 > 水晶1X + 水晶1.Width - 200 Then
        Command5.Visible = True
        over.Enabled = False
        inf "失败！"
    End If
    If X + Me.Width / 2 < 水晶1X + 水晶1.Width Then
        X = X + 150
    Else
        X = X - 150
    End If
ElseIf TotalHP2 <= 0 Then
    If X + Me.Width / 2 < 水晶2X + 水晶2.Width + 200 And X + Me.Width / 2 > 水晶2X + 水晶2.Width - 200 Then
        Command5.Visible = True
        over.Enabled = False
        inf "胜利！"
    End If
    If X + Me.Width / 2 < 水晶2X + 水晶2.Width Then
        X = X + 150
    Else
        X = X - 150
    End If
End If
    基地1.Left = 基地1X - X
    水晶1.Left = 水晶1X - X
    塔1.Left = 塔1X - X
    塔2.Left = 塔2X - X
    塔3.Left = 塔3X - X
    塔4.Left = 塔4X - X
    水晶2.Left = 水晶2X - X
    基地2.Left = 基地2X - X
    di.Left = -1000 'diX - X
    ziji.Left = -1000 'oldxx - X
    For a = 1 To 20
        If b11(a).e Then
            b1(a).Left = b11(a).X - X
        End If
        If b22(a).e Then
            b2(a).Left = b22(a).X - X
        End If
        If a <= 10 Then
            If daodan0(a).e Then
                daodan(a).Left = daodan0(a).X - X
            End If
        End If
    Next
End Sub

Private Sub Text3_DblClick()
'If MsgBox("是否关闭CPU刷新率限制。", vbYesNo) = vbYes Then
inf "系统：已恢复CPU刷新率上限。"
Text3.Text = 64
fpscpu = False
'End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If fpscpu = False Then
inf "已提高CPU刷新率上限，打破定时器限制。", 2
fpsshow_Click
fpscpu = True
End If
Select Case KeyAscii
Case 48 To 57, vbKeyBack 'Asc("0") To Asc("9") '允许0~9数字和退格键
Case Else
KeyAscii = 0
End Select

End Sub


Private Sub Timer1_Timer()
On Error GoTo err3:
If Rnd < 0.01 Then
Timer2.Interval = Int(1000 / (fps + 1) * 64) / 5
End If
DoEvents
If fpss2 <> Int(GetTickCount / 1000) Then
'MsgBox Int(GetTickCount / 1000)
'MsgBox fpss2
'MsgBox fpss
fps = fpss
fpss = 0
fpss2 = Int(GetTickCount / 1000)
End If
'fpsshow.Caption = "FPS:" & fpss

If IsNumeric(Text3) Then
If fpss / Int(Text3) > (GetTickCount - fpss2 * 1000) / 1000 Then Exit Sub
If fpss >= Int(Text3) And Int(Text3) > 20 Then
'fpsshow.Caption = "FPS:" & fpss & "ok"
Exit Sub
Else
fpss = fpss + 1

End If
End If
If Rnd < 0.2 Then fpsshow.Caption = "FPS:" & fps

If key = 1 Then X = X - 600: key = 0
If key = 2 Then X = X + 600: key = 0
If X < 0 Then X = 0: wantx = ""
toview X - (Me.Width - ziji.Width) / 2
xx xa1
Form1.Tag = X

'For a = 1 To 20
'If b11(a).e Then b11(a).x = b11(a).x + 50
'If b22(a).e Then b22(a).x = b22(a).x - 50
'Next
Exit Sub
err3:
inf "刷新帧时出现严重错误！——CH"
End Sub

Private Sub Timer2_Timer()
SetWindowPos Me.hwnd, -1, 0, 0, 0, 0, 3
If Rnd < 0.001 Then Form2.Show
Label1.Caption = "己方等级：" & lev1 & "|敌方等级:" & lev2 & "|敌方英雄血量" & diHP & "|敌方英雄蓝量" & diMP & "|用时" & sTot(Int(times)) & "s|自己已死" & TotalDied & "次！敌方已死" & TotalDied2 & "次！" & "己方水晶生命：" & TotalHP1 & "/5000|敌方水晶生命：" & TotalHP2 & "/5000"
If canq1 <= times And lev1 >= 2 Then qpic.Picture = LoadResPicture("Q1", vbResBitmap) Else qpic.Picture = LoadResPicture("Q2", vbResBitmap)
If canw1 <= times And lev1 >= 5 Then wpic.Picture = LoadResPicture("W1", vbResBitmap) Else wpic.Picture = LoadResPicture("W2", vbResBitmap)
If cane1 <= times And lev1 >= 8 Then epic.Picture = LoadResPicture("E1", vbResBitmap) Else epic.Picture = LoadResPicture("E2", vbResBitmap)
If canr1 <= times And lev1 >= 10 Then rpic.Picture = LoadResPicture("R1", vbResBitmap) Else rpic.Picture = LoadResPicture("R2", vbResBitmap)

If diHP < diMaxHP And diHP > 0 Then If Rnd < 0.2 Then diHP = diHP + 1
If diMP < diMaxMP And diMP > 0 Then If Rnd < 0.2 Then diMP = diMP + 1

If dieds = False Then
    If hp < Maxhp Then If Rnd < 0.2 Then c1 (hp + 1)
    If MP < Maxmp Then If Rnd < 0.2 Then c2 MP + 1
    
    If X < 3000 And X > -3000 Then
        c1 (hp + 160)
        c2 (MP + 160)
    End If
    If diHP > 0 Then
        If diX < 3000 + 基地2X And diX > -3000 + 基地2X Then
            If diHP <> diMaxHP Then diHP = diHP + 160
            If diHP > diMaxHP Then diHP = diMaxHP
            If diMP <> diMaxMP Then diMP = diMP + 160
            If diMP > diMaxMP Then diMP = diMaxMP
        End If
        If diHP = diMaxHP Then
            pcback = ""
        End If
    End If
End If
L1.Visible = False
L2.Visible = False
L3.Visible = False
L4.Visible = False

'fpsshow.Caption = "FPS:" & fps

If Format(times, "0.0") = "10.0" Then
inf "全军出击": reb = 5
End If

Dim ffv As Double
ffv = times
If ffv / 30 = Int(ffv / 30) And times <> 0 Then
inf "第" & Int(times / 30) & "轮出兵"
reb = 10
End If

If reb > 0 Then reb = reb - 1: ab1: ab2
If Int(times) = Int(oinf) Then infs.Visible = False
'If TotalHP1 < 0 Then End
'If TotalHP2 < 0 Then End

If tov = False Then updateview

Maxhp = 1000 + 100 * lev1
Maxmp = Maxhp
diMaxHP = 1000 + 100 * lev2
diMaxMP = diMaxHP
times = Format(times, "0.0") + 0.2
End Sub



Private Sub zuobi_Click()
Frame1.Visible = True
End Sub

Sub xx(xa)
'If xc <> 0 And xc <> "" Then
'xa1 = xa
'xb1 = xb
'xc1 = xc
'didi.Caption = "对象：" & xa & "，血：" & xb & "/" & xc
'didi2.Width = xb / xc * 3735
'Else
'
'End If
xa1 = xa
xx2 xa
End Sub
Sub xx2(xa)
Select Case xa
Case "基地1"
didi.Caption = "对象：基地1，血：100/100"
didi2.Width = 100 / 100 * 3735
Case "基地2"
didi.Caption = "对象：基地2，血：100/100"
didi2.Width = 100 / 100 * 3735
Case "水晶1"
If TotalHP1 <= 0 Then
didi.Caption = "对象：水晶1，死亡"
didi2.Width = 0
Else
didi.Caption = "对象：水晶1，血：" & TotalHP1 & "/5000"
didi2.Width = TotalHP1 / 5000 * 3735
End If
Case "水晶2"
If TotalHP2 <= 0 Then
didi.Caption = "对象：水晶2，死亡"
didi2.Width = 0
Else
didi.Caption = "对象：水晶2，血：" & TotalHP2 & "/5000"
didi2.Width = TotalHP2 / 5000 * 3735
End If
Case "塔1"
If T1hp <= 0 Then
didi.Caption = "对象：塔1，死亡"
didi2.Width = 0
Else
didi.Caption = "对象：塔1，血：" & T1hp & "/2000"
didi2.Width = T1hp / 2000 * 3735
End If
Case "塔2"
If T2hp <= 0 Then
didi.Caption = "对象：塔2，死亡"
didi2.Width = 0
Else
didi.Caption = "对象：塔2，血：" & T2hp & "/2000"
didi2.Width = T2hp / 2000 * 3735
End If
Case "塔3"
If T3hp <= 0 Then
didi.Caption = "对象：塔3，死亡"
didi2.Width = 0
Else
didi.Caption = "对象：塔3，血：" & T3hp & "/2000"
didi2.Width = T3hp / 2000 * 3735
End If
Case "塔4"
If T4hp <= 0 Then
didi.Caption = "对象：塔4，死亡"
didi2.Width = 0
Else
didi.Caption = "对象：塔4，血：" & T4hp & "/2000"
didi2.Width = T4hp / 2000 * 3735
End If
Case "敌人"
If diHP <= 0 Then
didi.Caption = "对象：敌人，死亡"
didi2.Width = 0
Else
didi.Caption = "对象：敌人，血：" & diHP & "/" & diMaxHP
didi2.Width = diHP / diMaxHP * 3735
End If
Case "小兵"
If b22(bingid).hp <= 0 Then
didi.Caption = "对象：小兵，死亡"
didi2.Width = 0
Else
didi.Caption = "对象：小兵" & bingid & "，血：" & b22(bingid).hp & "/500"
didi2.Width = b22(bingid).hp / 500 * 3735
End If
End Select
End Sub
Private Sub 基地1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
xx "基地1"
End Sub
Private Sub 水晶1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
xx "水晶1"
End Sub
Private Sub 水晶2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
xx "水晶2"
End Sub
Private Sub 基地2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
xx "基地2"
End Sub
Private Sub 塔1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
xx "塔1"
End Sub
Private Sub 塔2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
xx "塔2"
End Sub
Private Sub 塔3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
xx "塔3"
End Sub
Private Sub 塔4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
xx "塔4"
End Sub
Private Sub di_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
xx "敌人"
End Sub
