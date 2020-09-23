VERSION 5.00
Begin VB.Form frmSort 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sort"
   ClientHeight    =   1740
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5670
   Icon            =   "frmSort.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1740
   ScaleWidth      =   5670
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4320
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdSort 
      Caption         =   "&Sort"
      Height          =   375
      Left            =   4320
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.ComboBox cboSort 
      Height          =   315
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   840
      Width           =   3015
   End
   Begin VB.ComboBox cboField 
      Height          =   315
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   360
      Width           =   3015
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Sort Type:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sort in Field:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   360
      Width           =   975
   End
End
Attribute VB_Name = "frmSort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'File Name  : frmSort.frm
'Description: This will sort a recordset; ASCENDING
'             or DESCENDING.
'Author     : Masino Sinaga (masino_sinaga@yahoo.com)
'Web Site   : http://www30.brinkster.com/masinosinaga/
'             http://www.geocities.com/masino_sinaga/
'Date       : Sunday, November 3, 2002
'Location   : Puslatpos, Bandung 40151, INDONESIA
'-----------------------------------------------------

Private Sub cboField_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{Tab}"
End Sub
Private Sub cboSort_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{Tab}"
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

Private Sub cmdSort_Click()
Dim TipeSort As String
   If cboSort.Text = cboSort.List(0) Then
      TipeSort = "ASC"
   Else
      TipeSort = "DESC"
   End If
   Set adoSort = New ADODB.Recordset
   adoSort.Open "SHAPE " & _
     "{SELECT * FROM " & m_RecordSource & " " & _
     "ORDER BY " & cboField.Text & " " & TipeSort & "} " & _
     "AS ParentCMD APPEND " & _
     "({SELECT * FROM " & m_RecordSource & " " & _
     "ORDER BY " & cboField.Text & " " & TipeSort & "} " & _
     "AS ChildCMD RELATE " & m_FieldKey & " TO " & m_FieldKey & " ) " & _
     "AS ChildCMD", _
     db, adOpenStatic, adLockOptimistic
   With frmADOCode2
     On Error Resume Next
     If adoSort.RecordCount > 0 Then
        Set .grdDataGrid.DataSource = adoSort.DataSource
        Set .rsstrFindData = adoSort.DataSource
        Dim oTextData As TextBox
        For Each oTextData In .txtFields
            Set oTextData.DataSource = adoSort.DataSource
        Next
        Set .adoPrimaryRS = adoSort
        .cmdFirst.Value = True
     End If
   End With
   Exit Sub
Message:
     MsgBox Err.Number & " - " & _
            Err.Description, _
            vbExclamation, "No Result"

End Sub

Private Sub Form_Load()
On Error Resume Next
  If cboFind.Text = "" Then
     cmdFilter.Enabled = False
  Else
     cmdFilter.Enabled = True
  End If
  Set cnn = New ADODB.Connection
  cnn.ConnectionString = m_ConnectionString
  cnn.Open
  Set rs = New ADODB.Recordset
  rs.Open m_RecordSource, db, adOpenKeyset, adLockOptimistic, adCmdTable
  cboField.Clear
  For Each adoField In rs.Fields
      cboField.AddItem adoField.Name
  Next
  rs.Close
  cboField.Text = cboField.List(0)
  cboSort.AddItem "Ascending (ASC)"
  cboSort.AddItem "Descending (DESC)"
  cboSort.Text = cboSort.List(0)
  'Get setting for this form from INI File
  Call ReadFromINIToControls(frmSort, "Sort")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  'Save setting this form to INI File
  Call SaveFromControlsToINI(frmSort, "Sort")
  'Clear memory
  Set adoSort = Nothing
  Set adoField = Nothing
  Screen.MousePointer = vbDefault
  Unload Me
End Sub
