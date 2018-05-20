Attribute VB_Name = "modefix"
'Author: Zhang Yirong(all by zyr) Code name:Vientiane
'Dim button_help as sub in order to help user about button
Public intF_intR As Integer
Public Database As Connection
Public Declare Sub Sleep Lib "kernel32" _
                (ByVal dwMilliseconds As Long)
Public Sub setdatagird_width(DG As DataGrid, _
  adoData As ADODB.Recordset, intRecord As Integer, _
  intField As Integer, Optional accForHeaders As Boolean)
'DG = DataGrid(DataGirds)
'adoData = ADODB.Recordset(adomain_recordset)
'intRecord = Number of record(adomain_recordset.RecordCount)
'intField = Number of field(adomain_recordset.Fields.Count - 1)
'AccForHeaders = True or False(True)
Dim maxWidth As Single, width As Single, PB_schedule As Integer
Dim cellText As String, colkey As Long, superzyr As Long
Dim datagird_font As StdFont, datagird_scalemode As Integer
If intRecord = 0 Then Exit Sub
  Set datagird_font = DG.Parent.Font
  Set DG.Parent.Font = DG.Font
  datagird_scalemode = DG.Parent.ScaleMode
  DG.Parent.ScaleMode = vbTwips
  'object.ScaleMode = value(vbTwips)
  adoData.MoveFirst: maxWidth = 0
  'adoData is me dim not in environmental
  intF_intR = intField * intRecord
frmain.prgBar.Visible = True
frmain.prgBar.Max = intF_intR
For colkey = 0 To intField - 1
'for colkey or superzyr in order call data
  frmain.lblField.Caption = _
    "column:" & DG.Columns(colkey).DataField
  adoData.MoveFirst
  If accForHeaders = True Then
    maxWidth = DG.Parent.TextWidth(DG.Columns(colkey).Text) + 200
  End If: adoData.MoveFirst
  For superzyr = 0 To intRecord - 1
  'oh superzyr just me english so bad
    If intField <> 1 Then _
      cellText = DG.Columns(colkey).Text
    width = DG.Parent.TextWidth(cellText) + 200
    If width > maxWidth Then
      maxWidth = width: DG.Columns(colkey).width = maxWidth
    End If: adoData.MoveNext
    'Process next for in order choice width
    DoEvents
    PB_schedule = PB_schedule + 1
    frmain.lblAngka.Caption = _
      "finished:" & Format((PB_schedule _
        / intF_intR) * 100, "0") & "%"
    DoEvents
    frmain.prgBar.Value = PB_schedule
    DoEvents
  Next superzyr
DG.Columns(colkey).width = maxWidth
'oh ye set width in that just ok
Next colkey
Set DG.Parent.Font = datagird_font
DG.Parent.ScaleMode = datagird_scalemode
adoData.MoveFirst
'If finished move pointer to first record again
Sleep 100: frmain.prgBar.Value = 0: frmain.prgBar.Visible = False
frmain.lblAngka.Caption = "": frmain.lblField.Caption = ""
End Sub
Public Sub Database_connection()
'Core Database connection section
Set Database = New Connection
Database.CursorLocation = adUseClient
Database.Open "PROVIDER=MSDataShape;Data PROVIDER=" & _
      "Microsoft.Jet.OLEDB.4.0;Data Source=" _
            & App.Path & "\UserDB.mdb;"
End Sub


