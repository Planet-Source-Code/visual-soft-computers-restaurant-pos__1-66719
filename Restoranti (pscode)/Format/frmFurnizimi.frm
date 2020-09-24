VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmFurnizimi 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Delivery"
   ClientHeight    =   6975
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8895
   Icon            =   "frmFurnizimi.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   8895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtbar 
      Height          =   375
      Left            =   4560
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   6480
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txt1 
      Height          =   285
      Left            =   2520
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   6600
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtqmimi 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   5640
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox txtsasia 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7680
      TabIndex        =   8
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "POST"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      TabIndex        =   7
      Top             =   6360
      Width           =   1695
   End
   Begin Restauranti.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   120
      TabIndex        =   6
      Top             =   6240
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   53
   End
   Begin ComctlLib.ListView lst 
      Height          =   3855
      Left            =   120
      TabIndex        =   5
      Top             =   2160
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   6800
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "ID"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Description"
         Object.Width           =   6879
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Qty"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Price"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   4
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Total"
         Object.Width           =   1411
      EndProperty
   End
   Begin VB.ComboBox cboArt 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1680
      Width           =   4575
   End
   Begin Restauranti.ctrlLiner ctrlLiner2 
      Height          =   30
      Left            =   0
      TabIndex        =   1
      Top             =   1080
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   53
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Qty:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   3
      Left            =   7080
      TabIndex        =   11
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Price:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   5
      Left            =   4920
      TabIndex        =   10
      Top             =   1680
      Width           =   735
   End
   Begin ComctlLib.ImageList imgkerkimi 
      Left            =   5280
      Top             =   7320
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
            Picture         =   "frmFurnizimi.frx":058A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      Caption         =   "Select product:"
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
      TabIndex        =   4
      Top             =   1370
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   120
      Picture         =   "frmFurnizimi.frx":0844
      Top             =   120
      Width           =   810
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Delivery"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   1080
      TabIndex        =   3
      Top             =   120
      Width           =   4335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Select items from list and then click POST"
      Height          =   375
      Index           =   0
      Left            =   1080
      TabIndex        =   2
      Top             =   600
      Width           =   4335
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      BorderStyle     =   0  'Transparent
      Height          =   1095
      Left            =   0
      Top             =   0
      Width           =   10815
   End
End
Attribute VB_Name = "frmFurnizimi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lstid, lstpershkrimi, lstqmimi, lstsasia, lsttotal As String
Private Sub cboArt_Click()
Call dblidhja
With ar
criteria = "Select *From tblartikujt Where pershkrimi='" & cboArt & "'"
.Open criteria, strConek, adOpenStatic, adLockOptimistic
If .RecordCount >= 1 Then
txtbar = !id
txt1 = !pershkrimi
txtqmimi = Format(!qmimi, "###,###,##0.00")
txtsasia.SetFocus
Else
MsgBox "Item cannot found.", vbCritical, "Error!"
SendKeys "{end}+{home}"
Exit Sub
End If
.Close
End With
End Sub

Private Sub Command1_Click()
For ilst = 1 To lst.ListItems.Count
lstid = lst.ListItems(ilst).Text
lstpershkrimi = lst.ListItems(ilst).SubItems(1)
lstqmimi = lst.ListItems(ilst).SubItems(3)
lstsasia = lst.ListItems(ilst).SubItems(2)
lsttotal = lst.ListItems(ilst).SubItems(4)
'Ruajtja ne tabelen e Furnizimit
If lst.ListItems.Count = 0 Then
MsgBox "There is no items on delivery list", vbInformation, "Delivery"
Else
Set ac = New ADODB.Connection
Set ar = New ADODB.Recordset
Call dblidhja
ac.Open strConek
With ar
criteria = "Select *From tblFurnizimi"
.Open criteria, strConek, adOpenStatic, adLockOptimistic
.AddNew
!id = lstid
!pershkrimi = lstpershkrimi
!qmimi = lstqmimi
!sasia = lstsasia
!total = lsttotal
.Update
.Close
End With
lst.ListItems.Clear
txtsasia.Text = ""
txtqmimi.Text = ""
txt1.Text = ""
End If
'Ruajtja dhe Freskimi i stoqeve
Set ac = New ADODB.Connection
Set ar = New ADODB.Recordset
Call dblidhja
ac.Open strConek
With ar
criteria = "Select *From tblArtikujt Where id='" & lstid & "'"
.Open criteria, strConek, adOpenStatic, adLockOptimistic
!sasia = Val(!sasia) + lstsasia
.Update
.Close
End With
Next
End Sub

Private Sub Form_Load()
'Vendosja e artikujve në ComboBox
Call dblidhja
ar.Open "Select *From tblArtikujt", strConek, adOpenStatic, adLockOptimistic
ar.MoveFirst
Do While Not ar.EOF
cboArt.AddItem ar!pershkrimi
ar.MoveNext
Loop
ar.Close
End Sub
Private Sub lst_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyDelete
If lst.ListItems.Count = 0 Then
MsgBox "There is no items on delivery list", vbOKOnly, "Void item!"
Else
lst.ListItems.Remove (lst.SelectedItem.Index)
End If
End Select
End Sub
Private Sub txtsasia_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
lst.ListItems.Add , , txtbar, , 1
lst.ListItems(lst.ListItems.Count).SubItems(1) = txt1
lst.ListItems(lst.ListItems.Count).SubItems(2) = txtsasia
lst.ListItems(lst.ListItems.Count).SubItems(3) = txtqmimi
lst.ListItems(lst.ListItems.Count).SubItems(4) = Format(CCur(txtqmimi * txtsasia), "###,###,##0.00")
txtsasia.Text = ""
txtqmimi.Text = ""
txt1.Text = ""
cboArt.SetFocus
End If
End Sub
