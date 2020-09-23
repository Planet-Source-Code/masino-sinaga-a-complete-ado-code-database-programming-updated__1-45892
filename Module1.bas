Attribute VB_Name = "Module1"
'File Name  : Module1.bas
'Description: - Declare global variable
'             - Global Procedure/function
'Author     : Masino Sinaga (masino_sinaga@yahoo.com)
'Web Site   : http://www30.brinkster.com/masinosinaga/
'             http://www.geocities.com/masino_sinaga/
'Date       : Sunday, November 3, 20022
'Location   : Puslatpos, Bandung 40151, INDONESIA
'------------------------------------------------

Public db As Connection
Public adoBookMark As ADODB.Recordset
Public adoFind As ADODB.Recordset
Public adoFilter As ADODB.Recordset
Public adoSort As ADODB.Recordset
Public m_ConnectionString As String
Public m_RecordSource As String
Public m_FieldKey As String
Public maks As Integer

'Pause program for a while
Public Declare Sub Sleep Lib "kernel32" _
                         (ByVal dwMilliseconds As Long)

'Mencek setting program
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias _
"GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As _
Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, _
ByVal lpFileName As String) As Long

Public INIFileName As String

Public Declare Function WritePrivateProfileString Lib "kernel32" Alias _
"WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName _
As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long


Public Sub EnablePicture(InFrame As PictureBox, _
           ByVal Flag As Boolean)
'This will enabled or disable a frame with all controls
'in it
Dim Contrl As Control
On Error Resume Next
  InFrame.Enabled = Flag
  For Each Contrl In InFrame.Parent.Controls
     If (Contrl.Container.Name = InFrame.Name) Then
        If (TypeOf Contrl Is Frame) And Not _
           (Contrl.Name = InFrame.Name) Then
          EnablePicture Contrl, Flag
        Else
          If Not (TypeOf Contrl Is Menu) Then _
             Contrl.Enabled = Flag
        End If
     End If
  Next
End Sub

Public Sub Inisialisasi()
  m_ConnectionString = _
          "PROVIDER=MSDataShape;Data PROVIDER=" & _
          "Microsoft.Jet.OLEDB.4.0;Data Source=" _
          & App.Path & "\mahasiswa.mdb;"
  m_RecordSource = "t_mhs"
  m_FieldKey = "NIM"
End Sub

Public Sub UnlockTheFormKoneksi()
  Set db = New Connection
  db.CursorLocation = adUseClient
  'If you use database Access not protected by password
  'in the same location with application, you can use this.
  db.Open "PROVIDER=MSDataShape;Data PROVIDER=" & _
          "Microsoft.Jet.OLEDB.4.0;Data Source=" _
          & App.Path & "\mahasiswa.mdb;"
  
  'Database protected password...
  'db.Open "PROVIDER=MSDataShape;Data PROVIDER=" & _
  '        "Microsoft.Jet.OLEDB.4.0;Data Source=" _
  '        & App.Path & " \mahasiswa.mdb;Jet OLEDB:" & _
  '        "Database Password=passwordanda;"
  
  'If you use ODBC through DSN, here is the way...
  'db.Open "PROVIDER=MSDataShape;Data PROVIDER=MSDASQL;" & _
   '       "dsn=mahasiswa;uid=;pwd=;"

End Sub

Public Sub AdjustDataGridColumnWidth _
           (DG As DataGrid, _
           adoData As ADODB.Recordset, _
           intRecord As Integer, _
           intField As Integer, _
           Optional AccForHeaders As Boolean)

'This procedure will adjust DataGrids column width
'based on longest field in underlying source

'DG = DataGrid
'adoData = ADODB.Recordset
'intRecord = Number of record
'intField = Number of field
'AccForHeaders = True or False

    Dim row As Long, col As Long
    Dim width As Single, maxWidth As Single
    Dim saveFont As StdFont, saveScaleMode As Integer
    Dim cellText As String
    Dim i As Integer
    'If number of records = 0 then exit from the sub
    If intRecord = 0 Then Exit Sub
    'Save the form's font for DataGrid's font
    'We need this for form's TextWidth method
    Set saveFont = DG.Parent.Font
    Set DG.Parent.Font = DG.Font
    'Adjust ScaleMode to vbTwips for the form (parent).
    saveScaleMode = DG.Parent.ScaleMode
    DG.Parent.ScaleMode = vbTwips
    'Always from first record...
    adoData.MoveFirst
    maxWidth = 0
    
    'Tampung nilai maksimal dengan mengalikan variabel
    'jumlah field dan jumlah record, untuk menentukan
    'banyaknya sel yang akan diproses/disesuaikan lebarnya
    maks = intField * intRecord
    'Inisialisasi nilai maksimal progressbar
    frmADOCode2.prgBar1.Visible = True
    frmADOCode2.prgBar1.Max = maks
        
    'We begin from the first column until the last column
    For col = 0 To intField - 1
        'Tampilkan nama field/kolom yg sedang diproses
        frmADOCode2.lblField.Caption = _
           "Column: " & DG.Columns(col).DataField
        adoData.MoveFirst
        'Optional param, if true, set maxWidth to
        'width of DG.Parent
        If AccForHeaders Then
            maxWidth = DG.Parent.TextWidth(DG.Columns(col).Text) + 200
        End If
        'Repeat from first record again after we have
        'finished process the last record in
        'former column...
        adoData.MoveFirst
        For row = 0 To intRecord - 1
            'Get the text from the DataGrid's cell
            If intField = 1 Then
            Else  'If number of field more than one
                cellText = DG.Columns(col).Text
            End If
            'Fix the border...
            'Not for "multiple-line text"...
            width = DG.Parent.TextWidth(cellText) + 200
            'Update the maximum width if we found
            'the wider string...
            If width > maxWidth Then
               maxWidth = width
               DG.Columns(col).width = maxWidth
            End If
            'Process next record...
            adoData.MoveNext
            DoEvents
            'Counter bertambah satu, dst
            i = i + 1
            frmADOCode2.lblAngka.Caption = _
              "Finished " & Format((i / maks) * 100, "0") & "%"
            DoEvents
            frmADOCode2.prgBar1.Value = i
            DoEvents
            
        Next row
        'Change the column width...
        DG.Columns(col).width = maxWidth 'kolom terakhir!
    Next col
    'Change the DataGrid's parent property
    Set DG.Parent.Font = saveFont
    DG.Parent.ScaleMode = saveScaleMode
    'If finished, then move pointer to first record again
    adoData.MoveFirst
    Sleep 100
    ResetProgressBar
End Sub  'End of AdjustDataGridColumnWidth

Sub ResetProgressBar()
  With frmADOCode2
    .prgBar1.Value = 0
    .lblAngka.Caption = ""
    .lblField.Caption = ""
  End With
End Sub

Public Function SaveFromControlsToINI(Objek, MyAppName As String)
Dim Contrl As Control
Dim TempControlName As String, TempControlValue As String
On Error Resume Next
For Each Contrl In Objek
  If (TypeOf Contrl Is CheckBox) Or (TypeOf Contrl Is ComboBox) Then
    TempControlName = Contrl.Name
    TempControlValue = Contrl.Value
    If (TypeOf Contrl Is ComboBox) Then
      TempControlValue = Contrl.Text
      If TempControlValue = "" Then TempControlValue = 1
    End If
    Result = WritePrivateProfileString(MyAppName, TempControlName, _
    TempControlValue, INIFileName)
  End If
  If (TypeOf Contrl Is TextBox) Then
    TempControlName = Contrl.Name
    TempControlValue = Contrl.Text
    Result = WritePrivateProfileString(MyAppName, TempControlName, _
    TempControlValue, INIFileName)
  End If
  If (TypeOf Contrl Is OptionButton) Then
    TempControlValue = Contrl.Value
    If TempControlValue = True Then
      TempControlName = Contrl.Name
      TempControlValue = Contrl.Index
      Result = WritePrivateProfileString(MyAppName, TempControlName, _
      TempControlValue, INIFileName)
    End If
  End If
Next
End Function

Public Function ReadFromINIToControls(Objek, MyAppName As String)
Dim Contrl As Control
Dim TempControlName As String * 101, TempControlValue As String * 101
On Error Resume Next
For Each Contrl In Objek
If (TypeOf Contrl Is CheckBox) Or (TypeOf Contrl Is ComboBox) Or (TypeOf _
Contrl Is OptionButton) Or (TypeOf Contrl Is TextBox) Or (TypeOf Contrl Is CheckBox) Then
TempControlName = Contrl.Name
If (TypeOf Contrl Is TextBox) Or (TypeOf Contrl Is ComboBox) Then 'Or _
   '(TypeOf Contrl Is MaskEdBox) Then
   Result = GetPrivateProfileString(MyAppName, TempControlName, "", _
   TempControlValue, Len(TempControlValue), INIFileName)
Else 'If (TypeOf Contrl Is CheckBox) Then
   Result = GetPrivateProfileString(MyAppName, TempControlName, "0", _
   TempControlValue, Len(TempControlValue), INIFileName)
End If
  
If (TypeOf Contrl Is OptionButton) Then
   If Contrl.Index = Val(TempControlValue) Then Contrl = True
Else
    Contrl = TempControlValue
   If (TypeOf Contrl Is ComboBox) Then
      If Len(Contrl.Text) = 0 Then Contrl.ListIndex = 0
      End If
   End If
End If
Next
End Function
