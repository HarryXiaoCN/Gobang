VERSION 5.00
Begin VB.Form MF 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   Caption         =   "��Ȼ������"
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
   StartUpPosition =   2  '��Ļ����
   Begin VB.Timer ����ʱ�� 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   13080
      Top             =   8160
   End
   Begin VB.PictureBox ��� 
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
   Begin VB.PictureBox ��� 
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
   Begin VB.PictureBox ���� 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "΢���ź�"
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
      Begin VB.Line ����ʣ��ʱ����ʾ 
         BorderColor     =   &H00FFFF00&
         BorderWidth     =   5
         Visible         =   0   'False
         X1              =   3990
         X2              =   3990
         Y1              =   0
         Y2              =   7920
      End
      Begin VB.Label ʤ����ʾ 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         Caption         =   "��WIN!"
         BeginProperty Font 
            Name            =   "΢���ź�"
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
   Begin VB.Shape ������ʾ�� 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   5
      Height          =   2400
      Left            =   270
      Top             =   6150
      Visible         =   0   'False
      Width           =   2400
   End
   Begin VB.Menu ���̲˵� 
      Caption         =   "���̲˵�"
      Visible         =   0   'False
      Begin VB.Menu ��� 
         Caption         =   "���"
      End
      Begin VB.Menu ���� 
         Caption         =   "����"
      End
      Begin VB.Menu ���̲˵�cut1 
         Caption         =   "-"
      End
      Begin VB.Menu �弣 
         Caption         =   "�弣"
      End
   End
   Begin VB.Menu ��в˵�һ 
      Caption         =   "��в˵�"
      Visible         =   0   'False
      Begin VB.Menu ����һ 
         Caption         =   "����"
      End
      Begin VB.Menu ������ɫһ 
         Caption         =   "������ɫ"
      End
   End
   Begin VB.Menu ��в˵��� 
      Caption         =   "��в˵�"
      Visible         =   0   'False
      Begin VB.Menu ����� 
         Caption         =   "����"
      End
      Begin VB.Menu ������ɫ�� 
         Caption         =   "������ɫ"
      End
   End
End
Attribute VB_Name = "MF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private ִ����ɫ As Integer, ��ס������ As Long, ���ƶ����� As Boolean, ����������ɫ�� As Boolean, ��ͬ��ɫ As Integer
Private Type ����
    x As Single '�����ϵ�x����
    y As Single '�����ϵ�y����
    c As Integer 'ִ������/������ɫ������
End Type
Private Type ����
    x As Long
    y As Long
End Type
Private ���() As ����, ��������(7) As ����, �ֱ� As Single, �߱� As Single, ���ȱ� As Single
Private ʤ���ֱ� As Single, ʤ���߱� As Single, ���߱� As Single, ���߿� As Single
Private �ƶ�����ʱ����� As Single

Private Sub Form_Load()
    Dim i As Long
    Me.Caption = Me.Caption & " - Ver." & App.Major & "." & App.Minor & "." & App.Revision
    ��������(0).y = 1
    ��������(1).x = 1
    ��������(1).y = 1
    ��������(2).x = 1
    ��������(3).x = 1
    ��������(3).y = -1
    ��������(4).y = -1
    ��������(5).x = -1
    ��������(5).y = -1
    ��������(6).x = -1
    ��������(7).x = -1
    ��������(7).y = 1
    �ֱ� = ����.FontSize / ����.Width
    �߱� = ������ʾ��.BorderWidth / ���(1).Width
    ���ȱ� = ����ʣ��ʱ����ʾ.BorderWidth / ����.Width
    ʤ���ֱ� = ʤ����ʾ.FontSize / ����.Width
    ʤ���߱� = ʤ����ʾ.Height / ����.ScaleWidth
    ���߱� = 1 / ����.Width
    ���߿� = 1
    ReDim ���(0) '��ʼ����̬������֣�ʹ��ӵ��Ԫ�أ����(0)
    ����.Scale (0, 0)-(16, 16)
    ���̻���
    �ƶ�����ʱ����� = Timer()
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState <> 1 Then
        ����.Height = Me.Height - 1198
        ����.Width = ����.Height
        ����.Scale (0, 0)-(16, 16)
        ����.Left = Me.Width / 2 - ����.Width / 2
        
        ���(1).Left = ����.Left / 7
        ���(1).Width = ����.Left / 7 * 5
        ���(1).Height = ���(1).Width
        ���(1).Top = Me.Height - ���(1).Height - 835
        
        ���(2).Left = ����.Left + ����.Width + ����.Left / 7
        ���(2).Width = ���(1).Width
        ���(2).Height = ���(1).Height
        
        ������ʾ��.Width = ���(1).Width * 1.1
        ������ʾ��.Height = ������ʾ��.Width
        ������ʾ��.BorderWidth = �߱� * ���(1).Width
        
        Dim tmp As Single
        tmp = ���(1).Width * 0.05
        If ��ס������ = 1 Or ִ����ɫ = 1 Then
            ������ʾ��.Top = ���(1).Top - tmp
            ������ʾ��.Left = ���(1).Left - tmp
        ElseIf ��ס������ = 2 Or ִ����ɫ = 2 Then
            ������ʾ��.Top = ���(2).Top - tmp
            ������ʾ��.Left = ���(2).Left - tmp
        End If
        
        ����ʣ��ʱ����ʾ.BorderWidth = ���ȱ� * ����.Width
        ����ʣ��ʱ����ʾ.X1 = ����.ScaleHeight / 2
        ����ʣ��ʱ����ʾ.X2 = ����.ScaleWidth / 2
        If ����ʱ��.Enabled = False Then
            ����ʣ��ʱ����ʾ.Y1 = 0
            ����ʣ��ʱ����ʾ.Y2 = ����.ScaleHeight
        End If
        
        ʤ����ʾ.FontSize = ����.Width * ʤ���ֱ�
        ʤ����ʾ.Left = 0
        ʤ����ʾ.Width = ����.ScaleWidth
        ʤ����ʾ.Height = ����.ScaleWidth * ʤ���߱�
        ʤ����ʾ.Top = ����.ScaleHeight / 2 - ʤ����ʾ.Height / 2
        
        ���߿� = ���߱� * ����.Width
        
        ���̻���
    End If
End Sub

Private Sub ����ʱ��_Timer()
    If ����ʣ��ʱ����ʾ.Y1 <= 12 Then
        ����ʣ��ʱ����ʾ.Y1 = ����ʣ��ʱ����ʾ.Y1 + 6
        ����ʣ��ʱ����ʾ.Y2 = ����ʣ��ʱ����ʾ.Y2 + 6
    Else
        ֹͣ�������ӵȴ�
    End If
End Sub

Private Sub ������ɫһ_Click()
    ��ͬ��ɫ = 2
    ����������ɫ�� = True
    ����ʣ��ʱ����ʾ.Visible = True
    ����ʱ��.Enabled = True
End Sub

Private Sub ������ɫ��_Click()
    ��ͬ��ɫ = 1
    ����������ɫ�� = True
    ����ʣ��ʱ����ʾ.Visible = True
    ����ʱ��.Enabled = True
End Sub

Private Sub ֹͣ�������ӵȴ�()
    ����������ɫ�� = False
    ����ʱ��.Enabled = False
    ����ʣ��ʱ����ʾ.Visible = False
    ����ʣ��ʱ����ʾ.Y1 = 0
    ����ʣ��ʱ����ʾ.Y2 = 16
End Sub

Private Sub ���ӻغ�(����ID As Long)
    Dim ��ֻ���() As ����, i As Long
    ��ֻ��� = ���
    ReDim ���(UBound(���) - 1)
    For i = 1 To ����ID - 1
        ���(i) = ��ֻ���(i)
    Next
    For i = ����ID + 1 To UBound(��ֻ���)
        ���(i - 1) = ��ֻ���(i)
    Next
    ��ס������ = 0
    ���̻���
End Sub

Private Function �������Ӽ��(ByVal x As Long, ByVal y As Long) As Boolean
    Dim i As Long
    For i = 1 To UBound(���)
        With ���(i)
            If Int(.x + 0.5) = x And Int(.y + 0.5) = y And i <> ��ס������ Then
                '��鵽�����������������ӱ�ʶ���˳�����
                �������Ӽ�� = True
                Exit Function
            End If
        End With
    Next
End Function

Private Sub ���_DblClick(Index As Integer)
    If Index = 1 Then
        ����һ_Click
    Else
        �����_Click
    End If
End Sub

Private Sub ���_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        If ����������ɫ�� Then
            If ��ͬ��ɫ = Index Then
                Dim cT As Long
                cT = ���(��ͬ��ɫ).BackColor
                If ��ͬ��ɫ = 1 Then
                    ���(1).BackColor = ���(2).BackColor
                    ���(2).BackColor = cT
                Else
                    ���(2).BackColor = ���(1).BackColor
                    ���(1).BackColor = cT
                End If
                ֹͣ�������ӵȴ�
            End If
        Else
            If ��ס������ > 0 Then
                '�������������õ���
                ���ӻغ� ��ס������
                ������ʾ��.Visible = False
            ElseIf ִ����ɫ > 0 Then
                '����������ǵ���
                ִ����ɫ = 0
                ������ʾ��.Visible = False
            Else
                '����û������
                ִ����ɫ = Index
                ������ʾ��.Top = ���(Index).Top - ���(Index).Width * 0.05
                ������ʾ��.Left = ���(Index).Left - ���(Index).Width * 0.05
                ������ʾ��.Visible = True
            End If
        End If
    Else
        If Index = 1 Then
            PopupMenu ��в˵�һ
        Else
            PopupMenu ��в˵���
        End If
    End If
End Sub

Private Sub �弣_Click()
    �弣.Checked = Not �弣.Checked
    ���̻���
End Sub

Private Sub ����_DblClick()
    ��������
    ���̻���
End Sub

Private Sub ����_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        If ִ����ɫ > 0 And �������Ӽ��(Int(x + 0.5), Int(y + 0.5)) = False Then
            '��������,�������
            'ÿ��һ�ӣ������������һ��Ԫ���������������
            ReDim Preserve ���(UBound(���) + 1)
            With ���(UBound(���)) 'with�����ɲ���ÿ��ȡ��������ʱ��ȫ����
                .x = x '�ȼ��ڣ����(UBound(���)).x=x
                .y = y
                .c = ִ����ɫ
            End With
            ִ����ɫ = 0 '�������º������������
            ��ס������ = UBound(���) '�����µ�������Ϊ��ǰ��ס������
            ���ƶ����� = True
        ElseIf ��ס������ > 0 And �������Ӽ��(Int(x + 0.5), Int(y + 0.5)) = False Then
            '������ȡ��,�������
            With ���(��ס������)
                .x = x
                .y = y
            End With
            ��ס������ = 0
            ������ʾ��.Visible = False
        ElseIf ��ס������ = 0 And ִ����ɫ = 0 Then
            '��������
            ��ס������ = ��õ�������(x, y)
            ���ƶ����� = False
            If ��ס������ > 0 Then
                '������ס������ɫ���ı������ʾ��λ��
                ������ʾ��.Top = ���(���(��ס������).c).Top - ���(���(��ס������).c).Width * 0.05
                ������ʾ��.Left = ���(���(��ס������).c).Left - ���(���(��ס������).c).Width * 0.05
                ������ʾ��.Visible = True
            End If
        End If
        ���̻���
    Else
        PopupMenu ���̲˵�
    End If
End Sub

Private Sub ����_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 And ��ס������ > 0 And �������Ӽ��(Int(x + 0.5), Int(y + 0.5)) = False And Timer() - �ƶ�����ʱ����� > 0.01 Then
        '��ס���������ס����ʱ�������޸���ס���ӵ����굽������ڵ�λ���ϣ�����ƶ�
        ���ƶ����� = True
        With ���(��ס������)
            .x = x
            .y = y
        End With
        ���̻���
        �ƶ�����ʱ����� = Timer()
    End If
End Sub

Private Sub ����_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    'û�а�ס���ӿ�����Ϊ���뿪������
    If Button = 1 And ���ƶ����� = True Then
        ���ƶ����� = False
        ��ס������ = 0
        ������ʾ��.Visible = False
        ���̻���
    End If
End Sub

Private Function ��õ�������(x As Single, y As Single) As Long
    Dim i As Long
    For i = 1 To UBound(���)
        With ���(i)
            If x >= .x - 0.4 And x <= .x + 0.4 And y >= .y - 0.4 And y <= .y + 0.4 Then
                ��õ������� = i '����������i��ֵ
                Exit Function '�ҵ��������Ӻ�����������������ٱ�������������Ƿ����
            End If
        End With
    Next
End Function

Private Sub ���̻���()
    Dim i As Long
    
    If ����.Checked Then
        ��������
    End If
    
    ����.Cls '�����������
    
    '����������
    ����.DrawWidth = ���߿�
    ����.ForeColor = vbBlack
    ����.FontSize = ����.Width * �ֱ�
    For i = 1 To 15
        ����.Line (i, 1)-(i, 15)
        ����.Line (1, i)-(15, i)
        ����.CurrentX = 0
        ����.CurrentY = i - 0.4
        ����.Print i
        ����.CurrentX = i - 0.2
        ����.CurrentY = 0
        ����.Print Chr(64 + i)
    Next
    '�Ӵ����ܱ���
    ����.DrawWidth = ���߿� * 3
    ����.Line (1, 1)-(1, 15)
    ����.Line (15, 1)-(15, 15)
    ����.Line (1, 1)-(15, 1)
    ����.Line (1, 15)-(15, 15)
    ����.DrawWidth = ���߿�
    
    '���Ƹ�����
    ����.FillColor = vbBlack
    ����.Circle (4, 4), 0.1, vbBlack
    ����.Circle (12, 4), 0.1, vbBlack
    ����.Circle (4, 12), 0.1, vbBlack
    ����.Circle (12, 12), 0.1, vbBlack
    ����.Circle (8, 8), 0.1, vbBlack
    
    If ��ס������ > 0 Then
        ����.FillColor = vbRed
        ����.Circle (���(��ס������).x, ���(��ס������).y), 0.5, vbRed
    End If
    
    '������ּ�¼��������������������
    ����.FontSize = ����.Width * �ֱ� * 0.625
    For i = 1 To UBound(���)
        ����.FillColor = ���(���(i).c).BackColor
        ����.Circle (���(i).x, ���(i).y), 0.4, ���(���(i).c).BackColor
        If �弣.Checked Then
            ����.ForeColor = &H80000005 - ���(���(i).c).BackColor
            ����.CurrentX = ���(i).x - Len(Str(i)) / 9 + 0.07
            ����.CurrentY = ���(i).y - 0.25
            ����.Print i
        End If
    Next
    
    If UBound(���) > 8 Then
        ʤ�����
    End If
End Sub

Private Sub ʤ�����()
    Dim i As Long, v As Long, s As Long, ���̼��� As New Dictionary
    For i = 1 To UBound(���)
        ���̼���.Add Int(���(i).x + 0.5) & "," & Int(���(i).y + 0.5), ���(i).c
    Next
    For i = 1 To UBound(���)
        For v = 0 To 7
            s = ����ݹ�(���̼���, Int(���(i).x + 0.5), Int(���(i).y + 0.5), ���(i).c, v)
            If s >= 4 Then
                If ���(i).c = 1 Then
                    �����_Click
                Else
                    ����һ_Click
                End If
                Exit Sub
            End If
        Next
    Next
End Sub

Private Function ����ݹ�(d As Dictionary, x As Long, y As Long, c As Integer, v As Long) As Long
    Dim tmp As String
    tmp = x + ��������(v).x & "," & y + ��������(v).y
    If d.Exists(tmp) Then
        If d(tmp) = c Then
            ����ݹ� = ����ݹ�(d, x + ��������(v).x, y + ��������(v).y, c, v) + 1
            Exit Function
        End If
    End If
End Function

Private Sub ���_Click()
    ReDim ���(0)
    ���̻���
End Sub

Private Sub ����һ_Click()
    ʤ����ʾ.ForeColor = ���(2).BackColor
    ʤ����ʾ.Visible = True
End Sub

Private Sub �����_Click()
    ʤ����ʾ.ForeColor = ���(1).BackColor
    ʤ����ʾ.Visible = True
End Sub

Private Sub ʤ����ʾ_Click()
    ʤ����ʾ.Visible = False
End Sub

Private Sub ��������()
    Dim i As Long
    '��������
    If ִ����ɫ = 0 And ��ס������ = 0 Then
        For i = 1 To UBound(���)
            With ���(i)
                .x = Int(.x + 0.5)
                .y = Int(.y + 0.5)
            End With
        Next
    End If
End Sub
Private Sub ����_Click()
    ����.Checked = Not ����.Checked
    ���̻���
End Sub
