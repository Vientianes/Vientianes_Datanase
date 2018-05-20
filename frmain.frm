VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "By Vientianes"
   ClientHeight    =   4065
   ClientLeft      =   8295
   ClientTop       =   4035
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   6870
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picStatBoxReg 
      Height          =   600
      Left            =   60
      ScaleHeight     =   540
      ScaleWidth      =   5085
      TabIndex        =   27
      Top             =   3345
      Width           =   5145
      Begin VB.CommandButton cmdFirst 
         Caption         =   "first"
         Height          =   350
         Left            =   120
         TabIndex        =   15
         Top             =   100
         UseMaskColor    =   -1  'True
         Width           =   705
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "Prev"
         Height          =   350
         Left            =   840
         TabIndex        =   16
         Top             =   100
         UseMaskColor    =   -1  'True
         Width           =   705
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "Next"
         Height          =   350
         Left            =   3540
         TabIndex        =   17
         Top             =   90
         UseMaskColor    =   -1  'True
         Width           =   705
      End
      Begin VB.CommandButton cmdLast 
         Caption         =   "last"
         Height          =   350
         Left            =   4245
         TabIndex        =   18
         Top             =   90
         UseMaskColor    =   -1  'True
         Width           =   705
      End
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1500
         TabIndex        =   28
         Top             =   120
         Width           =   2115
      End
   End
   Begin VB.TextBox txtFields 
      DataField       =   "smain_remark"
      Height          =   285
      Index           =   5
      Left            =   3270
      MaxLength       =   8
      TabIndex        =   5
      Text            =   "5"
      Top             =   825
      Width           =   1695
   End
   Begin VB.TextBox txtFields 
      DataField       =   "smain_total"
      Height          =   285
      Index           =   4
      Left            =   690
      MaxLength       =   8
      TabIndex        =   4
      Text            =   "4"
      Top             =   825
      Width           =   1695
   End
   Begin VB.TextBox txtFields 
      DataField       =   "smain_price"
      Height          =   285
      Index           =   3
      Left            =   3090
      MaxLength       =   8
      TabIndex        =   3
      Text            =   "3"
      Top             =   465
      Width           =   1695
   End
   Begin VB.TextBox txtFields 
      DataField       =   "smian_unit"
      Height          =   285
      Index           =   2
      Left            =   600
      MaxLength       =   8
      TabIndex        =   2
      Text            =   "2"
      Top             =   465
      Width           =   1695
   End
   Begin VB.TextBox txtFields 
      DataField       =   "smain_name"
      Height          =   285
      Index           =   1
      Left            =   2880
      MaxLength       =   8
      TabIndex        =   1
      Text            =   "1"
      Top             =   90
      Width           =   1695
   End
   Begin VB.TextBox txtFields 
      DataField       =   "smain_id"
      Height          =   285
      Index           =   0
      Left            =   510
      MaxLength       =   8
      TabIndex        =   0
      Text            =   "0"
      Top             =   90
      Width           =   1695
   End
   Begin VB.PictureBox picButtons 
      Height          =   3630
      Left            =   5340
      ScaleHeight     =   3570
      ScaleWidth      =   1335
      TabIndex        =   19
      Top             =   195
      Width           =   1395
      Begin VB.CommandButton cmdAdd 
         Caption         =   "add new"
         Height          =   270
         Left            =   120
         TabIndex        =   6
         Top             =   150
         Width           =   1095
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "edit"
         Height          =   270
         Left            =   120
         TabIndex        =   10
         Top             =   1230
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "delete"
         Height          =   270
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   9
         Top             =   960
         Width           =   1095
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "refresh"
         Height          =   270
         Left            =   120
         TabIndex        =   11
         Top             =   1500
         Width           =   1095
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "close"
         Height          =   270
         Left            =   120
         TabIndex        =   13
         Top             =   3150
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "save"
         Height          =   270
         Left            =   120
         TabIndex        =   7
         Top             =   420
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "cancel"
         Height          =   270
         Left            =   120
         TabIndex        =   8
         Top             =   690
         Width           =   1095
      End
      Begin VB.CommandButton cmdDataGrid 
         Caption         =   "set width"
         Height          =   270
         Left            =   120
         TabIndex        =   12
         Top             =   2865
         Width           =   1095
      End
   End
   Begin ComctlLib.ProgressBar prgBar 
      Height          =   180
      Left            =   60
      TabIndex        =   26
      Top             =   3120
      Visible         =   0   'False
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   318
      _Version        =   327682
      Appearance      =   1
   End
   Begin MSDataGridLib.DataGrid DataGirds 
      Height          =   1680
      Left            =   120
      TabIndex        =   14
      Top             =   1185
      Width           =   4890
      _ExtentX        =   8625
      _ExtentY        =   2963
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label lblField 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   120
      TabIndex        =   30
      Top             =   2880
      Width           =   2655
   End
   Begin VB.Label lblAngka 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   3570
      TabIndex        =   29
      Top             =   2895
      Width           =   1455
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "name"
      Height          =   255
      Index           =   1
      Left            =   2400
      TabIndex        =   25
      Top             =   105
      Width           =   555
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "unit"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   24
      Top             =   495
      Width           =   555
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "price"
      Height          =   255
      Index           =   3
      Left            =   2475
      TabIndex        =   23
      Top             =   495
      Width           =   555
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "total"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   22
      Top             =   885
      Width           =   555
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "remarks"
      Height          =   255
      Index           =   5
      Left            =   2580
      TabIndex        =   21
      Top             =   855
      Width           =   555
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "IDs"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   20
      Top             =   120
      Width           =   315
   End
End
Attribute VB_Name = "frmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author: Zhang Yirong(all by zyr) Code name:Vientiane
'Dim button_help as sub in order to help user about button
Public WithEvents adomain_recordset As Recordset
Attribute adomain_recordset.VB_VarHelpID = -1
Dim Check_sameid As Recordset
Dim Save_DGplaces As Variant
Dim Edit_states As Boolean
Dim Add_states As Boolean
Private Sub cmdAdd_MouseMove(Button As Integer, _
  Shift As Integer, X As Single, Y As Single)
Call button_help("Add a record")
End Sub
Private Sub cmdUpdate_MouseMove(Button As Integer, _
  Shift As Integer, X As Single, Y As Single)
Call button_help("Save record")
End Sub
Private Sub cmdCancel_MouseMove(Button As Integer, _
  Shift As Integer, X As Single, Y As Single)
Call button_help("Cancel operation")
End Sub
Private Sub cmdDelete_MouseMove(Button As Integer, _
  Shift As Integer, X As Single, Y As Single)
Call button_help("Delete the selected record")
End Sub
Private Sub cmdEdit_MouseMove(Button As Integer, _
  Shift As Integer, X As Single, Y As Single)
Call button_help("Editor choice record")
End Sub
Private Sub cmdRefresh_MouseMove(Button As Integer, _
  Shift As Integer, X As Single, Y As Single)
Call button_help("Refresh the database and connection")
End Sub
Private Sub cmdDataGrid_MouseMove(Button As Integer, _
Shift As Integer, X As Single, Y As Single)
Call button_help("Set the width,i want see all")
End Sub
Private Sub cmdClose_MouseMove(Button As Integer, _
Shift As Integer, X As Single, Y As Single)
Call button_help("Sign out,check you saved?")
End Sub
Private Sub button_help(button_prompt As String)
frmain.Caption = button_prompt
End Sub
Private Sub txtFields_GotFocus(Index As Integer)
  txtFields(Index).BackColor = &HFFFF00
  txtFields(Index).SelStart = 0
  txtFields(Index).SelLength = Len(txtFields(Index))
End Sub
Private Sub txtFields_LostFocus(Index As Integer)
  txtFields(Index).BackColor = &H80000005
End Sub
Private Sub txtFields_KeyPress(Index As Integer, _
  KeyAscii As Integer)
Select Case Index
  Case 0 To 5
If KeyAscii = 13 Then SendKeys "{Tab}"
End Select
End Sub
Private Sub Form_Load()
Dim txtid As Integer
For txtid = 0 To 5
  txtFields(txtid) = ""
Next txtid
Call Database_connection
Call Controls_Setting(True, True)
Call Buttons_Setting(True)
DataGirds.Enabled = True
filenames = App.Path & "\Setting.ini"
Set adomain_recordset = New Recordset
adomain_recordset.Open "SHAPE {select smain_id,smain_name,smian_unit," & _
  "smain_price,smain_total,smain_remark from Stocks_main Order by smain_id} " & _
  "AS ParentCMD APPEND ({select smain_id," & _
  "smain_name,smian_unit,smain_price,smain_total,smain_remark FROM Stocks_main " & _
  "ORDER BY smain_id } AS ChildCMD RELATE smain_id " & _
  "TO smain_id) AS ChildCMD", Database, adOpenDynamic, adLockOptimistic
Dim Texts As TextBox
For Each Texts In Me.txtFields
  Set Texts.DataSource = adomain_recordset
Next
Set DataGirds.DataSource = adomain_recordset.DataSource
If adomain_recordset.RecordCount < 1 Then
  MsgBox "The record is empty!", vbExclamation, "Prompt"
End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If cmdUpdate.Enabled = True And cmdCancel.Enabled = True Then
  MsgBox "Please save or cancel" & vbCrLf, vbExclamation, "Prompt"
  cmdUpdate.SetFocus: Cancel = -1: Exit Sub
End If
If Not adomain_recordset Is Nothing Then _
  Set adomain_recordset = Nothing
txtFields(0).SetFocus
'Canael the focus of DataGird to prevent the exit error
Database.Close: Set Database = Nothing: End
End Sub
Private Sub Controls_Setting(Text_setting As Boolean, _
                         DataGird_setting As Boolean)
'Lock or unlock controls(txtFields)
Dim txtid As Integer
For txtid = 0 To 5
  txtFields(txtid).Locked = Text_setting
Next txtid: DataGirds.Enabled = DataGird_setting
End Sub
Private Sub Buttons_Setting(bVal As Boolean)
'Control of the button's enable property
cmdAdd.Enabled = bVal: cmdUpdate.Enabled = Not bVal
cmdCancel.Enabled = Not bVal: cmdDelete.Enabled = bVal
cmdEdit.Enabled = bVal: cmdRefresh.Enabled = bVal
cmdDataGrid.Enabled = bVal: cmdClose.Enabled = bVal
End Sub
Public Sub adomain_recordset_MoveComplete(ByVal adReason As _
      ADODB.EventReasonEnum, ByVal pError As _
      ADODB.Error, adStatus As ADODB.EventStatusEnum, _
      ByVal pRecordset As ADODB.Recordset)
lblStatus.Caption = Trim("Total/Now " & _
  CStr(adomain_recordset.AbsolutePosition) & _
  "/" & adomain_recordset.RecordCount)
With adomain_recordset
  If (.RecordCount > 1) Then
    If (.BOF) Or (.AbsolutePosition = 1) Then
      cmdFirst.Enabled = False: cmdPrevious.Enabled = False
      cmdNext.Enabled = True: cmdLast.Enabled = True
    ElseIf (.EOF) Or (.AbsolutePosition = .RecordCount) Then
      cmdNext.Enabled = False: cmdLast.Enabled = False
      cmdFirst.Enabled = True: cmdPrevious.Enabled = True
    Else
      cmdFirst.Enabled = True: cmdPrevious.Enabled = True
      cmdNext.Enabled = True: cmdLast.Enabled = True
    End If
  Else
    cmdFirst.Enabled = False: cmdPrevious.Enabled = False
    cmdNext.Enabled = False: cmdLast.Enabled = False
End If: End With: End Sub
Private Sub cmdAdd_Click()
On Error GoTo AddErr
With adomain_recordset
  If Not (.BOF And .EOF) Then _
    Save_DGplaces = .Bookmark
    Call Controls_Setting(False, False)
    .AddNew: Add_states = True
    lblStatus.Caption = "Adding a record "
    Call Buttons_Setting(False)
End With
On Error Resume Next
txtFields(0).SetFocus: Exit Sub
AddErr:
  MsgBox Err.Description & vbCrLf & _
  "find a errow,it is bad" & vbCrLf & _
  "please call zyr (QQ:429642909)"
End Sub
Private Sub cmdUpdate_Click()
Dim txtid As Integer
On Error GoTo UpdateErr
For txtid = 0 To 5
  If txtFields(txtid).Text = "" Then
    MsgBox "你必须填写所有的内容!", _
      vbExclamation, "Prompt"
    txtFields(txtid).SetFocus
Exit Sub: End If: Next txtid
Set Check_sameid = New Recordset
Check_sameid.Open "SELECT * FROM Stocks_main WHERE smain_id=" & _
  "'" & Trim(txtFields(0).Text) & "'", Database
If Check_sameid.RecordCount > 0 And Add_states Then
  MsgBox "smain_id '" & txtFields(0).Text & _
    "' already exist. " & vbCrLf & _
    "Please change to another smain_id!", _
    vbExclamation, "Double smain_id": txtFields(0).SetFocus
Set Check_sameid = Nothing: Exit Sub: End If
If Add_states Then adomain_recordset.MoveLast
If Edit_states Then adomain_recordset.MoveNext
Edit_states = False: Add_states = False
Call Buttons_Setting(True)
Call Controls_Setting(True, True)
lblStatus.Caption = "Record number " & _
  CStr(adomain_recordset.AbsolutePosition) & _
  " of " & adomain_recordset.RecordCount: Exit Sub
UpdateErr:
Select Case Err.Number
  Case -2147467259
  MsgBox "smain_id '" & txtFields(0).Text & _
    "' already exist." & vbCrLf & _
    "Please change to another smain_id!", _
    vbExclamation, "Double smain_id"
    txtFields(0).SetFocus
  Case Else
  MsgBox Err.Number & " - " & _
    Err.Description, vbCritical, "Error"
End Select: End Sub
Private Sub cmdCancel_Click()
On Error Resume Next
Call Controls_Setting(True, True)
Call Buttons_Setting(True)
Call cmdRefresh_Click
adomain_recordset.CancelUpdate
Edit_states = False: Add_states = False
If Save_DGplaces > 0 Then
  adomain_recordset.Bookmark = Save_DGplaces
Else
  adomain_recordset.MoveFirst
End If: End Sub
Private Sub cmdDelete_Click()
On Error GoTo DeleteErr
If MsgBox("你确定要删除吗？", vbQuestion + vbYesNo _
  + vbDefaultButton2) <> vbYes Then Exit Sub
adomain_recordset.Delete
adomain_recordset.MoveNext
If adomain_recordset.EOF Then _
 adomain_recordset.MoveLast
Exit Sub
DeleteErr:
  MsgBox Err.Description & vbCrLf & _
  "find a errow,it is bad" & vbCrLf & _
  "please call zyr(QQ:429642909)"
End Sub
Private Sub cmdEdit_Click()
On Error GoTo EditErr
lblStatus.Caption = "编辑记录"
Edit_states = True
With adomain_recordset
  If Not (.BOF And .EOF) Then _
    Save_DGplaces = .Bookmark
End With
Call Buttons_Setting(False)
Call Controls_Setting(False, False)
txtFields(0).SetFocus: Exit Sub
EditErr:
  MsgBox Err.Description & vbCrLf & _
  "find a errow,it is bad" & vbCrLf & _
  "please call zyr (QQ:429642909)"
End Sub
Private Sub cmdRefresh_Click()
On Error GoTo RefreshErr
Call Buttons_Setting(True)
Call Controls_Setting(True, True)
Set adofind_recordset = Nothing
Set DataGirds.DataSource = Nothing: Set adomain_recordset = New Recordset
adomain_recordset.Open "SHAPE {select smain_id,smain_name,smian_unit," & _
  "smain_price,smain_total,smain_remark from Stocks_main Order by smain_id} " & _
  "AS ParentCMD APPEND ({select smain_id," & _
  "smain_name,smian_unit,smain_price,smain_total,smain_remark FROM Stocks_main " & _
  "ORDER BY smain_id } AS ChildCMD RELATE smain_id " & _
  "TO smain_id) AS ChildCMD", Database, adOpenStatic, adLockOptimistic
Dim Texts As TextBox
For Each Texts In Me.txtFields
  Set Texts.DataSource = adomain_recordset
Next
Set DataGirds.DataSource = adomain_recordset.DataSource: Exit Sub
RefreshErr:
  MsgBox Err.Description & vbCrLf & _
  "find a errow,it is bad" & vbCrLf & _
  "please call zyr (QQ:429642909)"
Call Buttons_Setting(False): cmdRefresh.Enabled = True
cmdUpdate.Enabled = False: cmdCancel.Enabled = False
Edit_states = False: Add_states = False
adomain_recordset.CancelUpdate
If Save_DGplaces <> 0 Then
  adomain_recordset.Bookmark = Save_DGplaces
Else
  adomain_recordset.MoveFirst
End If: Exit Sub: End Sub
Private Sub cmdDataGrid_Click()
Dim intRecord As Integer
Dim intField As Integer
intRecord = adomain_recordset.RecordCount
intField = adomain_recordset.Fields.Count - 1
Call setdatagird_width(DataGirds, _
  adomain_recordset, intRecord, intField, True)
End Sub
Private Sub cmdClose_Click()
Unload Me
End Sub
Private Sub cmdFirst_Click()
On Error GoTo GoFirstError
adomain_recordset.MoveFirst: Exit Sub
GoFirstError:
  MsgBox Err.Description & vbCrLf & _
  "find a errow,it is bad" & vbCrLf & _
  "please call zyr (QQ:429642909)"
End Sub
Private Sub cmdPrevious_Click()
On Error GoTo GoPrevError
If Not adomain_recordset.BOF Then adomain_recordset.MovePrevious
  If adomain_recordset.BOF And adomain_recordset.RecordCount > 0 Then
    Beep
    adomain_recordset.MoveFirst
    MsgBox "This is the first record.", _
      vbInformation, "First Record"
  End If: Exit Sub
GoPrevError:
  MsgBox Err.Description & vbCrLf & _
  "find a errow,it is bad" & vbCrLf & _
  "please call zyr (QQ:429642909)"
End Sub
Private Sub cmdNext_Click()
On Error GoTo GoNextError
If Not adomain_recordset.EOF Then adomain_recordset.MoveNext
  If adomain_recordset.EOF And adomain_recordset.RecordCount > 0 Then
    Beep
    adomain_recordset.MoveLast
    MsgBox "This is the last record.", _
      vbInformation, "Last Record"
  End If: Exit Sub
GoNextError:
  MsgBox Err.Description & vbCrLf & _
  "find a errow,it is bad" & vbCrLf & _
  "please call zyr (QQ:429642909)"
End Sub
Private Sub cmdLast_Click()
On Error GoTo GoLastError
adomain_recordset.MoveLast: Exit Sub
GoLastError:
  MsgBox Err.Description & vbCrLf & _
  "find a errow,it is bad" & vbCrLf & _
  "please call zyr (QQ:429642909)"
End Sub
