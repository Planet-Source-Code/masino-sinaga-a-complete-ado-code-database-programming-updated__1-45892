VERSION 5.00
Begin VB.Form frmFind 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Find"
   ClientHeight    =   2055
   ClientLeft      =   3090
   ClientTop       =   6150
   ClientWidth     =   5700
   Icon            =   "frmFind.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   5700
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkMatch 
      Caption         =   "&Match whole word only"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1320
      Width           =   2175
   End
   Begin VB.CheckBox chkKonfirmasi 
      Caption         =   "&Display the complete data in found record"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Value           =   1  'Checked
      Width           =   3855
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4320
      TabIndex        =   6
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmdFindNext 
      Caption         =   "Find &Next"
      Height          =   375
      Left            =   4320
      TabIndex        =   5
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdFindFirst 
      Caption         =   "Find &First"
      Height          =   375
      Left            =   4320
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.ComboBox cboFind 
      Height          =   315
      Left            =   1200
      TabIndex        =   3
      Top             =   840
      Width           =   3015
   End
   Begin VB.ComboBox cboField 
      Height          =   315
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   360
      Width           =   3015
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Find what:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Find in Field:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   975
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'File Name  : frmFind.frm
'Description: Find data based on criteria from user.
'             I made my own procedure for searching data
'             based on (All Fields) or current selected
'             field; FindFirst and FindNext. As we know,
'             ADO don't have FindFirst and FindNext which
'             we can find in DAO. I hope these procedure
'             useful for you who used to use DAO before.
'Author     : Masino Sinaga (masino_sinaga@yahoo.com)
'Web Site   : http://www30.brinkster.com/masinosinaga/
'             http://www.geocities.com/masino_sinaga/
'Date       : Sunday, November 3, 2002
'Location   : Puslatpos, Bandung 40151, INDONESIA
'-----------------------------------------------------

Option Explicit

Dim cnn As ADODB.Connection
Dim WithEvents rs As ADODB.Recordset
Attribute rs.VB_VarHelpID = -1
Dim adoField As ADODB.Field
Dim mark As Variant
Dim intCount As Integer
Dim intPosition As Integer
Dim bFound As Boolean
Dim strFind As String, strFindNext As String
Dim strResult As String
Dim bCancel As Boolean


Private Sub cboField_Click()
  If cboField.Text = "(All Fields)" Then
     chkMatch.Value = 0
     chkMatch.Enabled = False
  Else
     chkMatch.Enabled = True
  End If
End Sub

Private Sub cboFind_Change()
  If Len(Trim(cboFind.Text)) > 0 Then
     cmdFindFirst.Enabled = True
     cmdFindFirst.Default = True
  Else
     cmdFindFirst.Enabled = False
     cmdFindNext.Enabled = False
  End If
End Sub

Private Sub cboFind_Click()
  If Len(Trim(cboFind.Text)) > 0 Then
     cmdFindFirst.Enabled = True
     cmdFindFirst.Default = True
  Else
     cmdFindFirst.Enabled = False
  End If
End Sub

Private Sub cboField_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     cboFind.SetFocus
     SendKeys "{Home}+{End}"
  End If
End Sub


Private Sub cmdFindFirst_Click()
Dim strFound As String
Dim i As Integer
'If criteria is not (All Fields)
If Trim(cboField.Text) <> "(All Fields)" Then
  On Error GoTo Message
  intCount = 0
  CheckDouble
  adoFind.MoveFirst
  bFound = False 'Not found yet
  Do While adoFind.EOF <> True
     DoEvents
     If bCancel = True Then 'If use interrupt by clicking
                            'Cancel button...
        Exit Sub            '... exit from this procedure
     End If
     
     If chkMatch.Value = 0 Then  'Not match whole word
       If InStr(UCase(adoFind.Fields(cboField.Text)), UCase(cboFind.Text)) > 0 Then
          DoEvents
          intCount = intCount + 1
          DoEvents
          'Get the absolute position
          intPosition = adoFind.AbsolutePosition
          'We found it, update bFound now
          bFound = True
       End If
     Else 'Match whole word only
       If UCase(adoFind.Fields(cboField.Text)) = UCase(cboFind.Text) Then
          DoEvents
          intCount = intCount + 1
          DoEvents
          'Get the absolute position
          intPosition = adoFind.AbsolutePosition
          'We found it, update bFound now
          bFound = True
       End If
     End If
     
     If intCount = 1 Then 'If this is the first found
        bFound = True 'Update bFound
        Exit Do       'Exit from this looping, because
                      'this is only the first time
     End If
     DoEvents
     adoFind.MoveNext
  Loop
  'Jika we found and intCount <> 0
  If bFound = True And intCount <> 0 Then
     'cmdFindNext ready
     cmdFindNext.Enabled = True
     'Display what position we found...
     strFound = "Found '" & cboFind.Text & _
                 "' in record number " & adoFind.AbsolutePosition & vbCrLf
     'This will get the name of field
     For i = 0 To adoFind.Fields.Count - 1
       'Get just field name that we need, but "ChildCMD"
       If adoFind.Fields(i).Name = "ChildCMD" Then
          Exit For
       End If
       'Get all data in record we found
       strFound = strFound & vbCrLf & _
            adoFind.Fields(i).Name & ": " & _
            vbTab & adoFind.Fields(i).Value
     Next i
  End If
  'If chkKonfirmasi was checked by user and data found
  If chkKonfirmasi.Value = 1 And bFound = True Then
     'Display in messagebox
     MsgBox strFound, vbInformation, "Found"
  End If
  If (adoFind.EOF) Then  'If pointer in end of recordset
     adoFind.MoveLast    'move to the last record
     bFound = False      'so, we haven't found it yet
     'Display messagebox we haven't found it
     MsgBox "'" & cboFind.Text & "' not found " & _
            "in field '" & cboField.Text & "'.", _
            vbExclamation, "Finished Searching"
     'cmdFindNext is not active because we haven't found
     'in cmdFindFirst
     cmdFindNext.Enabled = False
     Exit Sub
  End If
  Exit Sub
Else 'If user select (All Fields)
  FindFirstInAllFields '<-- call this procedure
  Exit Sub
End If
Message:
  MsgBox Err.Number & " - " & Err.Description
End Sub

Private Sub cmdFindNext_Click()
Dim strFound As String
Dim i As Integer
'If user select criteria: (All Fields)
If Trim(cboField.Text) <> "(All Fields)" Then
  On Error GoTo Message
  'First of all, we haven't found it, yet...
  bFound = False
  Do While adoFind.EOF <> True
     DoEvents
     If bCancel = True Then 'If use interrupt by clicking
                            'Cancel button...
        Exit Sub            '... exit from this procedure
     End If
     
     If chkMatch.Value = 0 Then  'Not match whole word
       'In FindNext, we compare the intPosition variable
       'with AbsolutePosition. If they are not same
       'then we found it
       If (InStr(UCase(adoFind.Fields(cboField.Text)), _
              UCase(cboFind.Text)) > 0) And _
              intPosition <> adoFind.AbsolutePosition Then
          DoEvents
          'Update counter position
          intCount = intCount + 1
          DoEvents
          'Get the absolute position
          intPosition = adoFind.AbsolutePosition
          'We found it, update bFound now
          bFound = True
       End If
     Else 'Match whole word only
       If UCase(adoFind.Fields(cboField.Text)) = _
              UCase(cboFind.Text) And _
              intPosition <> adoFind.AbsolutePosition Then
          DoEvents
          'Update counter position
          intCount = intCount + 1
          DoEvents
          'Get the absolute position
          intPosition = adoFind.AbsolutePosition
          'We found it, update bFound now
          bFound = True
       End If
     End If
     
     If bFound = True Then 'If we found it then
        Exit Do            'exit from this looping
     End If
     
     adoFind.MoveNext      'Process to next record
     DoEvents
     
     If adoFind.EOF Then   'If we are in EOF
        adoFind.MoveLast   'move to last record
        'Display message if we don't find it in looping
        MsgBox "'" & cboFind.Text & "' not found " & _
            "in field '" & cboField.Text & "'.", _
            vbExclamation, "Finished Searching"
        cmdFindNext.Enabled = False
        Exit Do
     End If
  Loop
  
  'If user check this checkbox
  If chkKonfirmasi.Value = 1 And _
     bFound = True And intCount <> 0 Then
     strFound = "Found '" & cboFind.Text & _
                 "' in record number " & adoFind.AbsolutePosition & vbCrLf
     'This iteration will get the name of all fields in
     'recordset, in order that we will display all data
     'in that record we found
     For i = 0 To adoFind.Fields.Count - 1
       'Check if the name contain "ChildCMD", exit from
       'iteration, we will not display this one.
       If adoFind.Fields(i).Name = "ChildCMD" Then
          Exit For
       End If
       'This will keep all data in record we found
       strFound = strFound & vbCrLf & _
            adoFind.Fields(i).Name & ": " & _
            vbTab & adoFind.Fields(i).Value
     Next i
     'Show the complete data in messagebox
     MsgBox strFound, vbInformation, "Found"
  End If

  If (adoFind.EOF) Then
     adoFind.MoveLast
     bFound = False 'We haven't found it, yet
     'Show messagebox
     MsgBox "'" & cboFind.Text & "' not found " & _
            "in field '" & cboField.Text & "'.", _
            vbExclamation, "Finished Searching"
     cmdFindNext.Enabled = False
     Exit Sub
  End If
  Exit Sub
Else 'If user select (All Fields)
  FindNextInAllFields  '<-- Call this procedure
  Exit Sub
End If
Message:
     adoFind.MoveLast
     bFound = False
     MsgBox "'" & cboFind.Text & "' not found " & _
            "in field '" & cboField.Text & "'.", _
            vbExclamation, "Finished Searching"
     cmdFindNext.Enabled = False
     Exit Sub
End Sub

Private Sub cmdCancel_Click()
  bCancel = True
  bFound = False
  'Empty string variable that we don't need
  m_ConnectionString = ""
  m_RecordSource = ""
  
  Set adoField = Nothing
  Set rs = Nothing
  Unload Me
  'Me.Hide  'Just hide this form, in order that we still
           'need the data later
End Sub

Private Sub Form_Load()
On Error Resume Next
  bCancel = False
  If cboFind.Text = "" Then
     cmdFindFirst.Enabled = False
     cmdFindNext.Enabled = False
  Else
     cmdFindFirst.Enabled = True
     cmdFindNext.Enabled = False
  End If
  Set cnn = New ADODB.Connection
  cnn.ConnectionString = m_ConnectionString
  cnn.Open
  Set rs = New ADODB.Recordset
  rs.Open m_RecordSource, db, adOpenKeyset, adLockOptimistic, adCmdTable
  cboField.Clear
  cboField.AddItem "(All Fields)"
  'This will get field name
  For Each adoField In rs.Fields
      cboField.AddItem adoField.Name
  Next
  rs.Close
  'Highlight the first item in combobox
  cboField.Text = cboField.List(0)
  'Get setting for this form from INI File
  Call ReadFromINIToControls(frmFind, "Find")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  'Save setting this form to INI File
  Call SaveFromControlsToINI(frmFind, "Find")
  'Clear memory
  Set adoFind = Nothing
  Set adoField = Nothing
  Screen.MousePointer = vbDefault
  Unload Me
End Sub

Private Sub CheckDouble()
Dim i As Integer
  If cboFind.Text = "" Then
     MsgBox "It can't not be a empty string!", _
            vbExclamation, "Invalid"
     cboFind.SetFocus
     Exit Sub
  End If
  For i = 0 To cboFind.ListCount - 1
    If cboFind.List(i) = cboFind.Text Then
       cboFind.SetFocus
       SendKeys "{Home}+{End}"
       Exit Sub
    End If
  Next i
  cboFind.AddItem cboFind.Text
  cboFind.Text = cboFind.List(cboFind.ListCount - 1)
End Sub

'This will search data in all fields for the very first time
Private Sub FindFirstInAllFields()
Dim strstrResult As String, strFound As String
Dim i As Integer, j As Integer, k As Integer
  'Always start from first record
  adoFind.MoveFirst
  strFind = cboFind.Text
  CheckDouble
Ulang:
  If adoFind.EOF And adoFind.RecordCount > 0 Then
     adoFind.MoveLast
     MsgBox "'" & strFind & "' not found in '" & cboField.Text & "'.", _
            vbExclamation, "Finished Searching"
     cmdFindNext.Enabled = False
     Exit Sub
  End If
  strstrResult = "":  strFound = ""
  With frmADOCode2
  For i = 0 To 4  'This iteration for data in textbox
      strResult = UCase(.txtFields(i).Text)
      If InStr(1, UCase(.txtFields(i).Text), UCase(strFind)) > 0 Then
         strstrResult = "" & strstrResult & "Found '" & strFind & "' at:" & vbCrLf & _
                      ""
       For j = 0 To 4 'This iteration for data in datagrid
          strResult = UCase(.txtFields(j).Text)
          If InStr(1, UCase(.txtFields(j).Text), UCase(strFind)) > 0 Then
             strFindNext = strFind
             'If we found it, tell user which position
             'it is...
              strstrResult = strstrResult & "" & vbCrLf & _
                 "  Record number " & CStr(adoFind.AbsolutePosition) & "" & vbCrLf & _
                 "  - Field name: " & .txtFields(j).DataField & "" & vbCrLf & _
                 "  - Contains: " & .txtFields(j).Text & "" & vbCrLf & _
                 "  - Columns number: " & j + 1 & " in DataGrid."
             For k = 0 To adoFind.Fields.Count - 1
                If adoFind.Fields(k).Name = "ChildCMD" Then
                  Exit For
               End If
               strFound = strFound & vbCrLf & _
                         adoFind.Fields(k).Name & ": " & _
                         vbTab & adoFind.Fields(k).Value
             Next k
             'Because we found, make cmdFindNext active...
             cmdFindNext.Enabled = True
             'If chkKonfirmasi was checked by user
             If chkKonfirmasi.Value = 1 Then
                'Display data
                 MsgBox strstrResult & vbCrLf & _
                        strFound, _
                        vbInformation, "Found"
             End If
          Else
          End If
       Next j  'End of iteration in datagrid
       Exit Sub
    Else
    End If
  Next i  'End of iteration in textBox
  End With
  'If we don't find in first record, move to next record
  adoFind.MoveNext
  GoTo Ulang
End Sub


'This will search data from the record position
'we found in FindFirstInAllFields procedure above.
'
Private Sub FindNextInAllFields()
Dim m As Integer, n As Integer, k As Integer
Dim strstrResult As String, strFound As String
strFindNext = strFind
If Len(Trim(strResult)) = 0 Then
   FindFirstInAllFields
   Exit Sub
End If
'Start from record position we found in FindFirstInAllFields
adoFind.MoveNext
strFound = "": strstrResult = ""
Ulang:
  'If we don't find it
  If adoFind.EOF And adoFind.RecordCount > 0 Then
     adoFind.MoveLast
     MsgBox "'" & strFindNext & "' not found in '" & cboField.Text & "'.", _
            vbExclamation, "Finished Searching"
     Exit Sub
  End If
  With frmADOCode2
  For n = 0 To 4  'This iteration for textbox
    strResult = UCase(cboFind.Text)
    'If we found it, all or similiar to it
    If InStr(1, UCase(.txtFields(n).Text), UCase(strFindNext)) > 0 Then
       strstrResult = "" & strstrResult & "Found '" & strFindNext & "' at:" & vbCrLf & _
                      ""
       For m = 0 To 4 'This iteration for datagrid
          strResult = UCase(cboFind.Text)
          If InStr(1, UCase(.txtFields(m).Text), UCase(strFindNext)) > 0 Then
             'If we found, tell user which record position
             'it is..
             strstrResult = strstrResult & vbCrLf & _
                 "  Record number " & CStr(adoFind.AbsolutePosition) & "" & vbCrLf & _
                 "  - Field name: " & .txtFields(m).DataField & "" & vbCrLf & _
                 "  - Contains: " & .txtFields(m).Text & "" & vbCrLf & _
                 "  - Column number: " & m + 1 & " in DataGrid."
             For k = 0 To adoFind.Fields.Count - 1
                If adoFind.Fields(k).Name = "ChildCMD" Then
                  Exit For
               End If
               'Get all data we found in that record
               strFound = strFound & vbCrLf & _
                         adoFind.Fields(k).Name & ": " & _
                         vbTab & adoFind.Fields(k).Value
             Next k
             'If chkKonfirmasi was checked by user
             If chkKonfirmasi.Value = 1 Then
                 'Display all data in that record we found
                 MsgBox strstrResult & vbCrLf & _
                        strFound, _
                        vbInformation, "Found"
                 cmdFindNext.Enabled = True
             End If
             Exit Sub
          Else
          End If
       Next m  'End of iteration in DataGrid
       Exit Sub
    Else
    End If
  Next n  'End of iteration in TextBox
  End With
  adoFind.MoveNext
  GoTo Ulang
End Sub
