VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmKerkoPerdoruesit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search Users"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5280
   Icon            =   "frmKerkoPerdoruesit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   5280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
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
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   4095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   375
      Left            =   3840
      TabIndex        =   3
      Top             =   3120
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "OK"
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   3120
      Width           =   1335
   End
   Begin Restauranti.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   120
      TabIndex        =   4
      Top             =   3000
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   53
   End
   Begin ComctlLib.ListView lv 
      Height          =   2175
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   3836
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      Icons           =   "ImageList2"
      SmallIcons      =   "ImageList2"
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Username"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Level"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Username:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   975
   End
   Begin ComctlLib.ImageList ImageList2 
      Left            =   6240
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   1
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmKerkoPerdoruesit.frx":058A
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FRMkERKOpERDORUESIT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
frmPerdoruesit.txtemri.Text = lv.SelectedItem.Text
Unload Me
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call dblidhja
With ar
criteria = "Select * From tblPerdoruesit Where emri like'" & Text1 & "%'"
.Open criteria, strConek, 3, 3
If .RecordCount >= 1 Then
lv.ListItems.Clear
Do While Not .EOF
lv.ListItems.Add , , !emri, , 1
lv.ListItems(lv.ListItems.Count).SubItems(1) = !niveli
.MoveNext
Loop
lv.SetFocus
Else
MsgBox "This record doesn't exist.", vbInformation, "Not found!"
Text1.SetFocus
Exit Sub
End If
.Close
End With
End If
End Sub
