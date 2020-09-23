VERSION 5.00
Begin VB.Form frmTest 
   BorderStyle     =   0  'None
   Caption         =   "SmoothForm"
   ClientHeight    =   6885
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6465
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmTest.frx":0000
   ScaleHeight     =   6885
   ScaleWidth      =   6465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdUnload 
      Caption         =   "Close"
      Height          =   375
      Left            =   4920
      TabIndex        =   3
      Top             =   6000
      Width           =   1095
   End
   Begin VB.HScrollBar scrlSmooth 
      Height          =   255
      LargeChange     =   10
      Left            =   1680
      Max             =   100
      TabIndex        =   0
      Top             =   360
      Value           =   25
      Width           =   1695
   End
   Begin VB.Label lbValue 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "25"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   1680
      TabIndex        =   2
      Top             =   600
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Curvature"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdUnload_Click()
    Unload Me
End Sub

Private Sub Form_Click()
    Unload Me
End Sub

'[Calling SmoothForm]
Private Sub Form_Load()
    SmoothForm Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim X As Integer
    If MsgBox("Is it Satisfactory?", vbQuestion + vbYesNo, "Please tell Me") = vbYes Then
        X = MsgBox("(  Please 'RATE' this code  ).Click 'Ok' to copy the site address  to your clipboard", vbInformation + vbOKCancel, "ThankYou")
    Else
        X = MsgBox("( Please give feedback ) to improve this code.Click 'Ok' to copy the site address  to your clipboard", vbInformation + vbOKCancel, "Please Give FeedBack")
    End If
    If X = vbOK Then Clipboard.SetText ("Not set")
End Sub

Private Sub scrlSmooth_Change()
    SmoothForm Me, (scrlSmooth + 1)
    lbValue = (scrlSmooth + 1)
End Sub
