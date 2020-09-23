VERSION 5.00
Begin VB.Form frmFilter 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Filter"
   ClientHeight    =   1740
   ClientLeft      =   3300
   ClientTop       =   6255
   ClientWidth     =   5670
   Icon            =   "frmFilter.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1740
   ScaleWidth      =   5670
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkMatch 
      Caption         =   "&Match whole word only"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   2175
   End
   Begin VB.ComboBox cboField 
      Height          =   315
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   360
      Width           =   3015
   End
   Begin VB.ComboBox cboFilter 
      Height          =   315
      Left            =   1200
      TabIndex        =   1
      Top             =   840
      Width           =   3015
   End
   Begin VB.CommandButton cmdFilter 
      Caption         =   "&Filter"
      Height          =   375
      Left            =   4320
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4320
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Filter in Field:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Filter What:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   975
   End
End
Attribute VB_Name = "frmFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'File Name  : frmFilter.frm
'Description: Filter a recordset based on value in
'             (All Fields)-> any fields in table...
'             or filter based on value in selected fields.
'Author     : Masino Sinaga (masino_sinaga@yahoo.com)
'Web Site   : http://www30.brinkster.com/masinosinaga/
'             http://www.geocities.com/masino_sinaga/
'Date       : Sunday, November 3, 2002
'Location   : Puslatpos, Bandung 40151, INDONESIA
'-----------------------------------------------------

Option Explicit        'All variables that we use
                       'must be declared

'The general variable that will be used in more
'than one procedure in this frmFilter module
Dim rs As ADODB.Recordset
Dim adoField As ADODB.Field
Dim cnn As ADODB.Connection
Dim DateField As Byte
Dim NumIntField As Byte
Dim NumDblField As Byte

Private Type arrTgl
  Name As String
  Tipe As String
End Type
Dim tabTgl() As arrTgl

Private Sub cboField_Click()
  If cboField.Text = "(All Fields)" Then
     chkMatch.Value = 0
     chkMatch.Enabled = False
  Else
     chkMatch.Enabled = True
  End If
End Sub

Private Sub cboField_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{Tab}"
End Sub
Private Sub cboFilter_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{Tab}"
End Sub

'If there is a change in cboFilter...
Private Sub cboFilter_Change()
  If Len(Trim(cboFilter.Text)) > 0 Then
     'cmdFilter will be active and ready
     cmdFilter.Enabled = True
     cmdFilter.Default = True
  Else 'Still empty
     cmdFilter.Enabled = False 'We can't use it
  End If
End Sub

Private Sub cmdFilter_Click()
On Error GoTo Message
 'Assign recordset variable to new recordset
  Set adoFilter = New ADODB.Recordset
 'Filter recordset based on paramter in SQL Statement
 AddCriteriaToCombo
 If cboField.Text <> "(All Fields)" Then
   If chkMatch.Value = 0 Then 'Not match whole criteria word
     adoFilter.Open "SHAPE " & _
     "{SELECT * FROM " & m_RecordSource & " " & _
     "WHERE " & Trim(cboField.Text) & " " & _
     "LIKE '%" & cboFilter.Text & "%' " & _
     "ORDER BY " & m_FieldKey & "} AS ParentCMD APPEND " & _
     "({SELECT * FROM " & m_RecordSource & " " & _
     "WHERE " & Trim(cboField.Text) & " " & _
     "LIKE '%" & cboFilter.Text & "%' " & _
     "ORDER BY " & m_FieldKey & "} AS ChildCMD " & _
     "RELATE " & m_FieldKey & " TO " & m_FieldKey & ") " & _
     "AS ChildCMD", db, adOpenStatic, adLockOptimistic
   Else 'Match whole criteria word only
     adoFilter.Open "SHAPE " & _
     "{SELECT * FROM " & m_RecordSource & " " & _
     "WHERE " & Trim(cboField.Text) & " " & _
     "= '%" & cboFilter.Text & "%' " & _
     "ORDER BY " & m_FieldKey & "} AS ParentCMD APPEND " & _
     "({SELECT * FROM " & m_RecordSource & " " & _
     "WHERE " & Trim(cboField.Text) & " " & _
     "= '%" & cboFilter.Text & "5' " & _
     "ORDER BY " & m_FieldKey & "} AS ChildCMD " & _
     "RELATE " & m_FieldKey & " TO " & m_FieldKey & ") " & _
     "AS ChildCMD", db, adOpenStatic, adLockOptimistic
   End If
   'Always reference to frmADOCode2 (main form in this example)
   With frmADOCode2
     'If recordset is not empty
     If adoFilter.RecordCount > 0 Then
       'Display the result to datagrid
       Set .grdDataGrid.DataSource = adoFilter.DataSource
       'This will update the status label in
       'middle of navigation button
       Set .rsstrFindData = adoFilter.DataSource
       'Bind the data to textbox
       Dim oTextData As TextBox
       For Each oTextData In .txtFields
           Set oTextData.DataSource = adoFilter.DataSource
       Next
       'Go to the first record
       .cmdFirst.Value = True
       'We can't click Bookmark button in order that
       'to prevent raise an error
       .cmdBookmark.Enabled = False
       'Update adoPrimaryRS in frmADOCode2 with
       'recordset that we have been filtered
       Set .adoPrimaryRS = adoFilter
     Else
       'If there is no recordset result
       .cmdRefresh.Value = True
       'Display message to user
       MsgBox "'" & cboFilter.Text & "' not found " & _
              "in field '" & cboField.Text & "'.", _
              vbExclamation, "No Result"
     End If
   End With
   Exit Sub
 Else
   FilterInAllFields
   Exit Sub
 End If
Message:
  MsgBox "'" & cboFilter.Text & "' not found " & _
         "in field '" & cboField.Text & "'.", _
         vbExclamation, "No Result"
End Sub

Private Sub cmdCancel_Click()
  'Empty string variable that we don't need
  m_ConnectionString = ""
  m_RecordSource = ""
  'Clear memory from object variable
  Set adoField = Nothing
  Set rs = Nothing
  Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
  'First, Filter button can not be accessed
  If cboFilter.Text = "" Then
     cmdFilter.Enabled = False
  Else 'If cboFilter is not empty
     cmdFilter.Enabled = True 'cmdFilter ready!
  End If
  Set cnn = New ADODB.Connection
  cnn.ConnectionString = m_ConnectionString
  cnn.Open
  Set rs = New ADODB.Recordset
  rs.Open m_RecordSource, db, adOpenKeyset, adLockOptimistic, adCmdTable
  cboField.Clear
  cboField.AddItem "(All Fields)"
  For Each adoField In rs.Fields
      cboField.AddItem adoField.Name
  Next
  'Highlight the first item in combobox
  cboField.Text = cboField.List(0)
  'Get setting for this form from INI File
  Call ReadFromINIToControls(frmFilter, "Filter")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  'Save setting this form to INI File
  Call SaveFromControlsToINI(frmFilter, "Filter")
  'Clear memory
  Set adoFilter = Nothing
  Set adoField = Nothing
  Screen.MousePointer = vbDefault
  Unload Me
End Sub

Private Sub FilterInAllFields()
Dim kriteria As String
Dim strCriteria As String, strField As String
Dim intField As Integer, i As Integer, j As Integer
Dim tabField() As String
'On Error GoTo Message
'On Error Resume Next
  'Always start from the first record
  rs.MoveFirst
  'To get the criteria and to make SQL Statement
  'shorter, we can use this way...
  strCriteria = ""
  intField = rs.Fields.Count
  ReDim tabField(intField)
  intField = rs.Fields.Count
  ReDim tabTgl(intField)
  'Dim i As Byte
  
  i = 0
  For Each adoField In rs.Fields
      tabField(i) = adoField.Name
      i = i + 1
  Next
      
  For i = 0 To intField - 1
    
    If chkMatch.Value = 0 Then 'Not match whole criteria word
  
     If i <> intField - 1 Then
        strField = strField & tabField(i) & ","
        strCriteria = strCriteria & _
           tabField(i) & " LIKE '%" & cboFilter.Text & "%' Or "
           
     Else
        strField = strField & tabField(i) & " "
        strCriteria = strCriteria & tabField(i) & " LIKE '%" & cboFilter.Text & "%' "
     End If
     
  Else  'Match whole criteria word only
    
     If i <> intField - 1 Then
        strField = strField & tabField(i) & ","
        
        strCriteria = strCriteria & _
           tabField(i) & " = '%" & cboFilter.Text & "%' Or "
           
     Else
        strField = strField & tabField(i) & " "
        strCriteria = strCriteria & tabField(i) & " = '%" & cboFilter.Text & "%' "
     End If
        
    
  End If
       
  Next i
  
  Set adoFilter = New ADODB.Recordset
     adoFilter.Open _
     "SHAPE " & _
     "{SELECT " & strField & " FROM " & m_RecordSource & " " & _
     "WHERE " & strCriteria & " ORDER BY " & m_FieldKey & "} " & _
     "AS ParentCMD APPEND " & _
     "({SELECT " & strField & " FROM " & m_RecordSource & " " & _
     "WHERE " & strCriteria & " ORDER BY " & m_FieldKey & "} " & _
     "AS ChildCMD RELATE " & m_FieldKey & " TO " & m_FieldKey & ") " & _
     "AS ChildCMD", db, adOpenStatic, adLockOptimistic
   
  With frmADOCode2
  If adoFilter.RecordCount > 0 Then
     Set .grdDataGrid.DataSource = adoFilter.DataSource
     Set .adoPrimaryRS = adoFilter
     Dim oTextData As TextBox
     For Each oTextData In .txtFields
         Set oTextData.DataSource = adoFilter.DataSource
     Next
     'Go to the first record, always
     .cmdFirst.Value = True
     'Bookmark button can not be accessed
     .cmdBookmark.Enabled = False
  Else
     .cmdRefresh.Value = True
     MsgBox "'" & cboFilter.Text & "' not found " & _
            "in field '" & cboField.Text & "'.", _
            vbExclamation, "No Result"
  End If
  End With
  Exit Sub
Message:
  'MsgBox Err.Number & " - " & Err.Description
  MsgBox "'" & cboFilter.Text & "' not found " & Chr(13) & _
         "in field '" & cboField.Text & "'.", _
         vbExclamation, "No Result"
End Sub

'This will check the double-criteria. If we find it,
'ignore it, else add it to the combobox...
Private Sub AddCriteriaToCombo()
Dim i As Integer
  If cboFilter.Text = "" Then
     MsgBox "Data is empty!", _
            vbExclamation, "Empty"
     cboFilter.SetFocus
     Exit Sub
  End If
  For i = 0 To cboFilter.ListCount - 1
    If cboFilter.List(i) = cboFilter.Text Then
       cboFilter.SetFocus
       SendKeys "{Home}+{End}"
       Exit Sub
    End If
  Next i
  cboFilter.AddItem cboFilter.Text
  cboFilter.Text = cboFilter.List(cboFilter.ListCount - 1)
End Sub
