VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3330
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   6000
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":0000
   ScaleHeight     =   3330
   ScaleWidth      =   6000
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1680
      Top             =   1680
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
On Error Resume Next
Unload Me
Dim Arq1 As String
Arq1 = ReadINI(App.Path & "\Restauranti.ini", "Fjalekalimi", "Fjalekalimi")
If Arq1 = 1 Then
frmLogin.Show 1
Else
frmMain.Show 1
End If
End Sub
