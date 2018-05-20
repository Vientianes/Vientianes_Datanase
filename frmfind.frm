VERSION 5.00
Begin VB.Form frmfind 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   2220
   ClientLeft      =   8655
   ClientTop       =   5130
   ClientWidth     =   6465
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   6465
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cboField 
      Height          =   300
      Left            =   1635
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   405
      Width           =   3015
   End
   Begin VB.ComboBox cboFind 
      Height          =   300
      Left            =   1635
      TabIndex        =   1
      Top             =   885
      Width           =   3015
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "查找"
      Height          =   375
      Left            =   4755
      TabIndex        =   2
      Top             =   165
      Width           =   1215
   End
   Begin VB.CommandButton cmdFinds 
      Caption         =   "查找下一个"
      Height          =   375
      Left            =   4755
      TabIndex        =   3
      Top             =   645
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消"
      Height          =   375
      Left            =   4755
      TabIndex        =   4
      Top             =   1245
      Width           =   1215
   End
   Begin VB.CheckBox Cek_allrecords 
      Caption         =   "显示全部记录"
      Height          =   255
      Left            =   555
      TabIndex        =   6
      Top             =   1725
      Value           =   1  'Checked
      Width           =   3855
   End
   Begin VB.CheckBox Cek_script 
      Caption         =   "查找文字"
      Height          =   255
      Left            =   525
      TabIndex        =   5
      Top             =   1365
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "要查找的字段"
      Height          =   255
      Left            =   465
      TabIndex        =   8
      Top             =   405
      Width           =   1125
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "查找的内容"
      Height          =   255
      Left            =   465
      TabIndex        =   7
      Top             =   885
      Width           =   1125
   End
End
Attribute VB_Name = "frmfind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim adofield As ADODB.Field
Dim adocon As ADODB.Connection
Dim WithEvents adorec As ADODB.Recordset
Attribute adorec.VB_VarHelpID = -1
Private Sub cmdCancel_Click()
'Just hide this form, in order
'that we still need the data later
Set adofield = Nothing
Set adorec = Nothing
Unload Me: End Sub
Private Sub Form_Load()
On Error Resume Next
If cboField.Text = "" Then
  cmdFind.Enabled = False
  cmdFinds.Enabled = False
End If
Set adocon = New ADODB.Connection
adocon.ConnectionString = _
  "PROVIDER=MSDataShape;Data PROVIDER=" & _
  "Microsoft.Jet.OLEDB.4.0;Data Source=" _
  & App.Path & "\UserDB.mdb;": adocon.Open
Set adorec = New ADODB.Recordset
adorec.Open "Stocks_main", Database, _
  adOpenKeyset, adLockOptimistic, adCmdTable
cboField.Clear: cboField.AddItem "All fields"
For Each adofield In adorec.fields
  cboField.AddItem adofield.Name
Next: adorec.Clone
cboField.Text = cboField.List(0)
End Sub
Private Sub Form_QueryUnload(Cancel _
  As Integer, UnloadMode As Integer)
Set adofind_recordset = Nothing
Set adofield = Nothing: Unload Me: End Sub
