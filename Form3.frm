VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   ".test.PPT.run"
   ClientHeight    =   1170
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   2475
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1170
   ScaleWidth      =   2475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "检测到PPT正在运行！"
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
      Begin VB.CommandButton Command1 
         Caption         =   "隐藏倒计时30min"
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1935
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
