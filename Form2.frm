VERSION 5.00
Begin VB.Form Form2 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Òþ²Ø15·ÖÖÓ£¿"
   ClientHeight    =   810
   ClientLeft      =   13230
   ClientTop       =   5535
   ClientWidth     =   1530
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   810
   ScaleWidth      =   1530
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer3 
      Interval        =   3000
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer Timer2 
      Left            =   1080
      Top             =   0
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Òþ²Ø15·ÖÖÓ"
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1575
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Timer2.Enabled = True
Timer3.Enabled = False
Form1.Visible = False
Form2.Visible = False
End Sub

Private Sub Form_Load()
Timer2.Enabled = False
Timer3.Enabled = True
Timer3.Interval = 3000
End Sub

Private Sub Timer2_Timer()
Static s_Minutes     As Long
            
          s_Minutes = s_Minutes + 1
          If s_Minutes = 15 Then
                  s_Minutes = 0
                  Form1.Visible = True
                  Form2.Visible = False
          End If
  End Sub


Private Sub Timer3_Timer()
Form1.Visible = True
Form2.Visible = False
End Sub
