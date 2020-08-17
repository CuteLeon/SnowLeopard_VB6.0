VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1920
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1920
   Icon            =   "小雪豹.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   1920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Timer1 
      Interval        =   75
      Left            =   1080
      Top             =   1080
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Dim PicturePath As String

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Dim scrPT As POINTAPI

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        GetCursorPos scrPT
        Timer1.Enabled = False
        MakeTrans Me, App.Path & "\图片资源\dials_17.PNG"
    Else
        End
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Me.WindowState <> 2 And Button = 1 Then
        Dim pt As POINTAPI
        MakeTrans Me, App.Path & "\图片资源\dials_18.PNG"
        GetCursorPos pt
        Me.Left = Me.Left + (pt.X - scrPT.X) * 15
        Me.Top = Me.Top + (pt.Y - scrPT.Y) * 15
        scrPT = pt
    End If
End Sub

Private Sub Form_Load()
    PicturePath = App.Path & "\图片资源\dials_0.PNG"
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
    Static PicIndex As Integer
    PicturePath = App.Path & "\图片资源\dials_" & PicIndex & ".PNG"
    SetWindowPos Me.hWnd, -1, 0, 0, 0, 0, &H10 Or &H40 Or &H2 Or &H1
    If Dir(PicturePath) <> "" Then
        MakeTrans Me, PicturePath
        PicIndex = PicIndex + 1
    Else
        PicIndex = 0
    End If
End Sub
