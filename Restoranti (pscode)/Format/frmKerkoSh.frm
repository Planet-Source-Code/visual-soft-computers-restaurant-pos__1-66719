VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmKerkoKat 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search Category"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4440
   Icon            =   "frmKerkoSh.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   4440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.ListView lvkerkimi 
      Height          =   1935
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   3413
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      Icons           =   "imgkerkimi"
      SmallIcons      =   "imgkerkimi"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Category"
         Object.Width           =   5292
      EndProperty
   End
   Begin VB.TextBox txtKerko 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   120
      TabIndex        =   0
      Top             =   420
      Width           =   4215
   End
   Begin ComctlLib.ImageList imgkerkimi 
      Left            =   2400
      Top             =   3840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   14
      ImageHeight     =   14
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   1
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmKerkoSh.frx":058A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Category name:"
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
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "frmKerkoKat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
End Sub

Private Sub lvkerkimi_Click()
If lvkerkimi.ListItems.Count = 0 Then
Exit Sub
Else
frmKategoria.Text1.Text = lvkerkimi.ListItems(lvkerkimi.SelectedItem.Index).Text
Unload Me
End If
End Sub

Private Sub txtKerko_KeyPress(KeyAscii As Integer)
On Error Resume Next
Dim inti As ListItem
If KeyAscii = 13 Then
Set ac = New ADODB.Connection
Set ar = New ADODB.Recordset
Call dblidhja
ac.Open strConek
lvkerkimi.ListItems.Clear
With ar
criteria = "Select *From tblkategoria"
.Open criteria, strConek, adOpenStatic, adLockOptimistic
.MoveFirst
Do While Not .EOF
If Mid(!kategoria, 1, Len(txtKerkimi)) = txtKerkimi Then
Set intitem = lvkerkimi.ListItems.Add(, , !kategoria, , 1)
End If
.MoveNext
lvkerkimi.SetFocus
Loop
.Close
End With
End If
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
