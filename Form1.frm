VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "��߿�"
   ClientHeight    =   690
   ClientLeft      =   17985
   ClientTop       =   315
   ClientWidth     =   915
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   690
   ScaleWidth      =   915
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Left            =   -240
      Top             =   -240
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   645
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   1035
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Days As Integer
Dim Times As Date
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST& = -1
' �����������б�������λ���κ�������ڵ�ǰ��
Private Const SWP_NOSIZE& = &H1
' ���ִ��ڴ�С
Private Const SWP_NOMOVE& = &H2
' ���ִ���λ��


Function FindProcess(ProcessName) As Boolean
Dim ps
'ö�ٽ���
For Each ps In GetObject("winmgmts:\\.\root\cimv2:win32_process").instances_ 'ѭ������
If UCase(ps.Name) = UCase(ProcessName) Then
   FindProcess = True
   Exit Function
End If
Next
End Function


Private Sub Form_Load()
'Label2.Visible = False
Dim Days As Integer
Dim Times As Date
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
' ��������Ϊ������ǰ
Timer1.Enabled = True
Timer1.Interval = 500
If FindProcess("POWERPNT.EXE") Then
Form2.Visible = True
End If
End Sub


Private Sub Label1_Click()
Form2.Visible = True
End Sub

Private Sub Timer1_Timer()
Const date0 As Date = #6/7/2018#        '10:10:10 AM# ����
Days = CInt(date0 - Now - 0.5)  'ȡ��������(����)
Times = CDate(date0 - Now)
Label1.Caption = Days
End Sub
