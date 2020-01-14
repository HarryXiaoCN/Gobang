VERSION 5.00
Begin VB.Form MF 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   Caption         =   "自然五子棋"
   ClientHeight    =   8685
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13665
   DrawWidth       =   3
   Icon            =   "MF.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Picture         =   "MF.frx":6932
   ScaleHeight     =   8685
   ScaleWidth      =   13665
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer 交换时钟 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   13080
      Top             =   8160
   End
   Begin VB.PictureBox 棋盒 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2000
      Index           =   2
      Left            =   11280
      ScaleHeight     =   1965
      ScaleWidth      =   1965
      TabIndex        =   2
      Top             =   360
      Width           =   2000
   End
   Begin VB.PictureBox 棋盒 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   2000
      Index           =   1
      Left            =   480
      ScaleHeight     =   1965
      ScaleWidth      =   1965
      TabIndex        =   1
      Top             =   6360
      Width           =   2000
   End
   Begin VB.PictureBox 棋盘 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   8000
      Left            =   2880
      ScaleHeight     =   7965
      ScaleWidth      =   7965
      TabIndex        =   0
      Top             =   360
      Width           =   8000
      Begin VB.Line 交换剩余时间提示 
         BorderColor     =   &H00FFFF00&
         BorderWidth     =   5
         Visible         =   0   'False
         X1              =   3990
         X2              =   3990
         Y1              =   0
         Y2              =   7920
      End
      Begin VB.Label 胜利提示 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         Caption         =   "●WIN!"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   42
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1125
         Left            =   0
         TabIndex        =   3
         Top             =   3240
         Visible         =   0   'False
         Width           =   7995
      End
   End
   Begin VB.Shape 持子提示框 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   5
      Height          =   2400
      Left            =   270
      Top             =   6150
      Visible         =   0   'False
      Width           =   2400
   End
   Begin VB.Menu 棋盘菜单 
      Caption         =   "棋盘菜单"
      Visible         =   0   'False
      Begin VB.Menu 清空 
         Caption         =   "清空"
      End
      Begin VB.Menu 整理 
         Caption         =   "整理"
      End
      Begin VB.Menu 棋盘菜单cut1 
         Caption         =   "-"
      End
      Begin VB.Menu 棋迹 
         Caption         =   "棋迹"
      End
   End
   Begin VB.Menu 棋盒菜单一 
      Caption         =   "棋盒菜单"
      Visible         =   0   'False
      Begin VB.Menu 认输一 
         Caption         =   "认输"
      End
      Begin VB.Menu 交换颜色一 
         Caption         =   "交换颜色"
      End
   End
   Begin VB.Menu 棋盒菜单二 
      Caption         =   "棋盒菜单"
      Visible         =   0   'False
      Begin VB.Menu 认输二 
         Caption         =   "认输"
      End
      Begin VB.Menu 交换颜色二 
         Caption         =   "交换颜色"
      End
   End
End
Attribute VB_Name = "MF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private 执棋颜色 As Integer, 摁住的棋子 As Long, 在移动棋子 As Boolean, 交换棋子颜色中 As Boolean, 需同意色 As Integer
Private Type 棋子
    x As Single '棋盘上的x坐标
    y As Single '棋盘上的y坐标
    c As Integer '执棋类型/棋子颜色的索引
End Type
Private Type 向量
    x As Long
    y As Long
End Type
Private 棋局() As 棋子, 方向向量(7) As 向量, 字比 As Single, 线比 As Single, 进度比 As Single
Private 胜利字比 As Single, 胜利高比 As Single, 网线比 As Single, 网线宽 As Single
Private 移动绘制时间记忆 As Single

Private Sub Form_Load()
    Dim i As Long
    Me.Caption = Me.Caption & " - Ver." & App.Major & "." & App.Minor & "." & App.Revision
    方向向量(0).y = 1
    方向向量(1).x = 1
    方向向量(1).y = 1
    方向向量(2).x = 1
    方向向量(3).x = 1
    方向向量(3).y = -1
    方向向量(4).y = -1
    方向向量(5).x = -1
    方向向量(5).y = -1
    方向向量(6).x = -1
    方向向量(7).x = -1
    方向向量(7).y = 1
    字比 = 棋盘.FontSize / 棋盘.Width
    线比 = 持子提示框.BorderWidth / 棋盒(1).Width
    进度比 = 交换剩余时间提示.BorderWidth / 棋盘.Width
    胜利字比 = 胜利提示.FontSize / 棋盘.Width
    胜利高比 = 胜利提示.Height / 棋盘.ScaleWidth
    网线比 = 1 / 棋盘.Width
    网线宽 = 1
    ReDim 棋局(0) '初始化动态数组棋局，使其拥有元素：棋局(0)
    棋盘.Scale (0, 0)-(16, 16)
    棋盘绘制
    移动绘制时间记忆 = Timer()
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState <> 1 Then
        棋盘.Height = Me.Height - 1198
        棋盘.Width = 棋盘.Height
        棋盘.Scale (0, 0)-(16, 16)
        棋盘.Left = Me.Width / 2 - 棋盘.Width / 2
        
        棋盒(1).Left = 棋盘.Left / 7
        棋盒(1).Width = 棋盘.Left / 7 * 5
        棋盒(1).Height = 棋盒(1).Width
        棋盒(1).Top = Me.Height - 棋盒(1).Height - 835
        
        棋盒(2).Left = 棋盘.Left + 棋盘.Width + 棋盘.Left / 7
        棋盒(2).Width = 棋盒(1).Width
        棋盒(2).Height = 棋盒(1).Height
        
        持子提示框.Width = 棋盒(1).Width * 1.1
        持子提示框.Height = 持子提示框.Width
        持子提示框.BorderWidth = 线比 * 棋盒(1).Width
        
        Dim tmp As Single
        tmp = 棋盒(1).Width * 0.05
        If 摁住的棋子 = 1 Or 执棋颜色 = 1 Then
            持子提示框.Top = 棋盒(1).Top - tmp
            持子提示框.Left = 棋盒(1).Left - tmp
        ElseIf 摁住的棋子 = 2 Or 执棋颜色 = 2 Then
            持子提示框.Top = 棋盒(2).Top - tmp
            持子提示框.Left = 棋盒(2).Left - tmp
        End If
        
        交换剩余时间提示.BorderWidth = 进度比 * 棋盘.Width
        交换剩余时间提示.X1 = 棋盘.ScaleHeight / 2
        交换剩余时间提示.X2 = 棋盘.ScaleWidth / 2
        If 交换时钟.Enabled = False Then
            交换剩余时间提示.Y1 = 0
            交换剩余时间提示.Y2 = 棋盘.ScaleHeight
        End If
        
        胜利提示.FontSize = 棋盘.Width * 胜利字比
        胜利提示.Left = 0
        胜利提示.Width = 棋盘.ScaleWidth
        胜利提示.Height = 棋盘.ScaleWidth * 胜利高比
        胜利提示.Top = 棋盘.ScaleHeight / 2 - 胜利提示.Height / 2
        
        网线宽 = 网线比 * 棋盘.Width
        
        棋盘绘制
    End If
End Sub

Private Sub 交换时钟_Timer()
    If 交换剩余时间提示.Y1 <= 12 Then
        交换剩余时间提示.Y1 = 交换剩余时间提示.Y1 + 6
        交换剩余时间提示.Y2 = 交换剩余时间提示.Y2 + 6
    Else
        停止交换棋子等待
    End If
End Sub

Private Sub 交换颜色一_Click()
    需同意色 = 2
    交换棋子颜色中 = True
    交换剩余时间提示.Visible = True
    交换时钟.Enabled = True
End Sub

Private Sub 交换颜色二_Click()
    需同意色 = 1
    交换棋子颜色中 = True
    交换剩余时间提示.Visible = True
    交换时钟.Enabled = True
End Sub

Private Sub 停止交换棋子等待()
    交换棋子颜色中 = False
    交换时钟.Enabled = False
    交换剩余时间提示.Visible = False
    交换剩余时间提示.Y1 = 0
    交换剩余时间提示.Y2 = 16
End Sub

Private Sub 棋子回盒(棋子ID As Long)
    Dim 棋局缓存() As 棋子, i As Long
    棋局缓存 = 棋局
    ReDim 棋局(UBound(棋局) - 1)
    For i = 1 To 棋子ID - 1
        棋局(i) = 棋局缓存(i)
    Next
    For i = 棋子ID + 1 To UBound(棋局缓存)
        棋局(i - 1) = 棋局缓存(i)
    Next
    摁住的棋子 = 0
    棋盘绘制
End Sub

Private Function 已有棋子检查(ByVal x As Long, ByVal y As Long) As Boolean
    Dim i As Long
    For i = 1 To UBound(棋局)
        With 棋局(i)
            If Int(.x + 0.5) = x And Int(.y + 0.5) = y And i <> 摁住的棋子 Then
                '检查到有棋子立马反馈该棋子标识并退出函数
                已有棋子检查 = True
                Exit Function
            End If
        End With
    Next
End Function

Private Sub 棋盒_DblClick(Index As Integer)
    If Index = 1 Then
        认输一_Click
    Else
        认输二_Click
    End If
End Sub

Private Sub 棋盒_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        If 交换棋子颜色中 Then
            If 需同意色 = Index Then
                Dim cT As Long
                cT = 棋盒(需同意色).BackColor
                If 需同意色 = 1 Then
                    棋盒(1).BackColor = 棋盒(2).BackColor
                    棋盒(2).BackColor = cT
                Else
                    棋盒(2).BackColor = 棋盒(1).BackColor
                    棋盒(1).BackColor = cT
                End If
                停止交换棋子等待
            End If
        Else
            If 摁住的棋子 > 0 Then
                '手上有棋盘上拿的子
                棋子回盒 摁住的棋子
                持子提示框.Visible = False
            ElseIf 执棋颜色 > 0 Then
                '手上有棋盒那的子
                执棋颜色 = 0
                持子提示框.Visible = False
            Else
                '手上没有棋子
                执棋颜色 = Index
                持子提示框.Top = 棋盒(Index).Top - 棋盒(Index).Width * 0.05
                持子提示框.Left = 棋盒(Index).Left - 棋盒(Index).Width * 0.05
                持子提示框.Visible = True
            End If
        End If
    Else
        If Index = 1 Then
            PopupMenu 棋盒菜单一
        Else
            PopupMenu 棋盒菜单二
        End If
    End If
End Sub

Private Sub 棋迹_Click()
    棋迹.Checked = Not 棋迹.Checked
    棋盘绘制
End Sub

Private Sub 棋盘_DblClick()
    整理棋盘
    棋盘绘制
End Sub

Private Sub 棋盘_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        If 执棋颜色 > 0 And 已有棋子检查(Int(x + 0.5), Int(y + 0.5)) = False Then
            '手中有子,落点无子
            '每落一子，棋局数组增加一个元素用来存放新棋子
            ReDim Preserve 棋局(UBound(棋局) + 1)
            With 棋局(UBound(棋局)) 'with方法可不比每次取用类属性时键全类名
                .x = x '等价于：棋局(UBound(棋局)).x=x
                .y = y
                .c = 执棋颜色
            End With
            执棋颜色 = 0 '棋子落下后手上棋子清空
            摁住的棋子 = UBound(棋局) '将落下的棋子作为当前摁住的棋子
            在移动棋子 = True
        ElseIf 摁住的棋子 > 0 And 已有棋子检查(Int(x + 0.5), Int(y + 0.5)) = False Then
            '手中有取子,落点无子
            With 棋局(摁住的棋子)
                .x = x
                .y = y
            End With
            摁住的棋子 = 0
            持子提示框.Visible = False
        ElseIf 摁住的棋子 = 0 And 执棋颜色 = 0 Then
            '手中无子
            摁住的棋子 = 获得点上棋子(x, y)
            在移动棋子 = False
            If 摁住的棋子 > 0 Then
                '根据摁住棋子颜色，改变持子提示框位置
                持子提示框.Top = 棋盒(棋局(摁住的棋子).c).Top - 棋盒(棋局(摁住的棋子).c).Width * 0.05
                持子提示框.Left = 棋盒(棋局(摁住的棋子).c).Left - 棋盒(棋局(摁住的棋子).c).Width * 0.05
                持子提示框.Visible = True
            End If
        End If
        棋盘绘制
    Else
        PopupMenu 棋盘菜单
    End If
End Sub

Private Sub 棋盘_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 And 摁住的棋子 > 0 And 已有棋子检查(Int(x + 0.5), Int(y + 0.5)) = False And Timer() - 移动绘制时间记忆 > 0.01 Then
        '按住鼠标且有摁住棋子时，不断修改摁住棋子的坐标到鼠标现在的位置上，造成移动
        在移动棋子 = True
        With 棋局(摁住的棋子)
            .x = x
            .y = y
        End With
        棋盘绘制
        移动绘制时间记忆 = Timer()
    End If
End Sub

Private Sub 棋盘_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    '没有按住棋子咯，因为手离开了棋盘
    If Button = 1 And 在移动棋子 = True Then
        在移动棋子 = False
        摁住的棋子 = 0
        持子提示框.Visible = False
        棋盘绘制
    End If
End Sub

Private Function 获得点上棋子(x As Single, y As Single) As Long
    Dim i As Long
    For i = 1 To UBound(棋局)
        With 棋局(i)
            If x >= .x - 0.4 And x <= .x + 0.4 And y >= .y - 0.4 And y <= .y + 0.4 Then
                获得点上棋子 = i '函数将返回i的值
                Exit Function '找到点上棋子后立马结束函数，不再遍历后面的棋子是否符合
            End If
        End With
    Next
End Function

Private Sub 棋盘绘制()
    Dim i As Long
    
    If 整理.Checked Then
        整理棋盘
    End If
    
    棋盘.Cls '清空棋盘内容
    
    '绘制棋盘线
    棋盘.DrawWidth = 网线宽
    棋盘.ForeColor = vbBlack
    棋盘.FontSize = 棋盘.Width * 字比
    For i = 1 To 15
        棋盘.Line (i, 1)-(i, 15)
        棋盘.Line (1, i)-(15, i)
        棋盘.CurrentX = 0
        棋盘.CurrentY = i - 0.4
        棋盘.Print i
        棋盘.CurrentX = i - 0.2
        棋盘.CurrentY = 0
        棋盘.Print Chr(64 + i)
    Next
    '加粗四周边线
    棋盘.DrawWidth = 网线宽 * 3
    棋盘.Line (1, 1)-(1, 15)
    棋盘.Line (15, 1)-(15, 15)
    棋盘.Line (1, 1)-(15, 1)
    棋盘.Line (1, 15)-(15, 15)
    棋盘.DrawWidth = 网线宽
    
    '绘制辅助点
    棋盘.FillColor = vbBlack
    棋盘.Circle (4, 4), 0.1, vbBlack
    棋盘.Circle (12, 4), 0.1, vbBlack
    棋盘.Circle (4, 12), 0.1, vbBlack
    棋盘.Circle (12, 12), 0.1, vbBlack
    棋盘.Circle (8, 8), 0.1, vbBlack
    
    If 摁住的棋子 > 0 Then
        棋盘.FillColor = vbRed
        棋盘.Circle (棋局(摁住的棋子).x, 棋局(摁住的棋子).y), 0.5, vbRed
    End If
    
    '根据棋局记录的棋子属性来绘制棋子
    棋盘.FontSize = 棋盘.Width * 字比 * 0.625
    For i = 1 To UBound(棋局)
        棋盘.FillColor = 棋盒(棋局(i).c).BackColor
        棋盘.Circle (棋局(i).x, 棋局(i).y), 0.4, 棋盒(棋局(i).c).BackColor
        If 棋迹.Checked Then
            棋盘.ForeColor = &H80000005 - 棋盒(棋局(i).c).BackColor
            棋盘.CurrentX = 棋局(i).x - Len(Str(i)) / 9 + 0.07
            棋盘.CurrentY = 棋局(i).y - 0.25
            棋盘.Print i
        End If
    Next
    
    If UBound(棋局) > 8 Then
        胜负检查
    End If
End Sub

Private Sub 胜负检查()
    Dim i As Long, v As Long, s As Long, 棋盘记忆 As New Dictionary
    For i = 1 To UBound(棋局)
        棋盘记忆.Add Int(棋局(i).x + 0.5) & "," & Int(棋局(i).y + 0.5), 棋局(i).c
    Next
    For i = 1 To UBound(棋局)
        For v = 0 To 7
            s = 方向递归(棋盘记忆, Int(棋局(i).x + 0.5), Int(棋局(i).y + 0.5), 棋局(i).c, v)
            If s >= 4 Then
                If 棋局(i).c = 1 Then
                    认输二_Click
                Else
                    认输一_Click
                End If
                Exit Sub
            End If
        Next
    Next
End Sub

Private Function 方向递归(d As Dictionary, x As Long, y As Long, c As Integer, v As Long) As Long
    Dim tmp As String
    tmp = x + 方向向量(v).x & "," & y + 方向向量(v).y
    If d.Exists(tmp) Then
        If d(tmp) = c Then
            方向递归 = 方向递归(d, x + 方向向量(v).x, y + 方向向量(v).y, c, v) + 1
            Exit Function
        End If
    End If
End Function

Private Sub 清空_Click()
    ReDim 棋局(0)
    棋盘绘制
End Sub

Private Sub 认输一_Click()
    胜利提示.ForeColor = 棋盒(2).BackColor
    胜利提示.Visible = True
End Sub

Private Sub 认输二_Click()
    胜利提示.ForeColor = 棋盒(1).BackColor
    胜利提示.Visible = True
End Sub

Private Sub 胜利提示_Click()
    胜利提示.Visible = False
End Sub

Private Sub 整理棋盘()
    Dim i As Long
    '规整棋盘
    If 执棋颜色 = 0 And 摁住的棋子 = 0 Then
        For i = 1 To UBound(棋局)
            With 棋局(i)
                .x = Int(.x + 0.5)
                .y = Int(.y + 0.5)
            End With
        Next
    End If
End Sub
Private Sub 整理_Click()
    整理.Checked = Not 整理.Checked
    棋盘绘制
End Sub
