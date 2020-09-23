VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmADOCode2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Demo ADO Code Database Programming (c) Masino Sinaga, November 3, 2002"
   ClientHeight    =   6510
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   7995
   Icon            =   "frmADOCode2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   7995
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.ProgressBar prgBar1 
      Height          =   180
      Left            =   120
      TabIndex        =   41
      Top             =   5280
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   318
      _Version        =   327682
      Appearance      =   1
   End
   Begin MSMask.MaskEdBox mskTglLahir 
      Height          =   285
      Left            =   1320
      TabIndex        =   4
      Top             =   1455
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.ComboBox cboStatus 
      Height          =   315
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   2050
      Width           =   1335
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Status"
      Height          =   285
      Index           =   6
      Left            =   5880
      MaxLength       =   1
      TabIndex        =   38
      Top             =   2090
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Sex"
      Height          =   285
      Index           =   5
      Left            =   5880
      MaxLength       =   1
      TabIndex        =   37
      Top             =   1770
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picStatBoxReg 
      Height          =   600
      Left            =   120
      ScaleHeight     =   540
      ScaleWidth      =   6075
      TabIndex        =   34
      Top             =   5520
      Width           =   6135
      Begin VB.CommandButton cmdLast 
         Caption         =   "Last"
         Height          =   350
         Left            =   5280
         TabIndex        =   25
         Top             =   100
         UseMaskColor    =   -1  'True
         Width           =   705
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "Next"
         Height          =   350
         Left            =   4560
         TabIndex        =   24
         Top             =   100
         UseMaskColor    =   -1  'True
         Width           =   705
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "Prev"
         Height          =   350
         Left            =   840
         TabIndex        =   23
         Top             =   100
         UseMaskColor    =   -1  'True
         Width           =   705
      End
      Begin VB.CommandButton cmdFirst 
         Caption         =   "First"
         Height          =   350
         Left            =   120
         TabIndex        =   22
         Top             =   100
         UseMaskColor    =   -1  'True
         Width           =   705
      End
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1440
         TabIndex        =   35
         Top             =   120
         Width           =   3240
      End
   End
   Begin MSDataGridLib.DataGrid grdDataGrid 
      Height          =   2500
      Left            =   120
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   2520
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   4419
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
   Begin VB.PictureBox picButtons 
      Height          =   5985
      Left            =   6480
      ScaleHeight     =   5925
      ScaleWidth      =   1335
      TabIndex        =   33
      Top             =   120
      Width           =   1395
      Begin VB.CommandButton cmdDataGrid 
         Caption         =   "Data&Grid"
         Height          =   350
         Left            =   120
         TabIndex        =   19
         Top             =   4080
         Width           =   1095
      End
      Begin VB.CommandButton cmdBookmark 
         Caption         =   "&Bookmark..."
         Height          =   350
         Left            =   120
         TabIndex        =   18
         Top             =   3720
         Width           =   1095
      End
      Begin VB.CommandButton cmdFilter 
         Caption         =   "&Filter..."
         Height          =   350
         Left            =   120
         TabIndex        =   15
         Top             =   2640
         Width           =   1095
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "Fi&nd..."
         Height          =   350
         Left            =   120
         TabIndex        =   14
         Top             =   2280
         Width           =   1095
      End
      Begin VB.CommandButton cmdUnFilter 
         Caption         =   "&Unfilter"
         Height          =   350
         Left            =   120
         TabIndex        =   16
         Top             =   3000
         Width           =   1095
      End
      Begin VB.CommandButton cmdSort 
         Caption         =   "S&ort..."
         Height          =   350
         Left            =   120
         TabIndex        =   17
         Top             =   3360
         Width           =   1095
      End
      Begin VB.CommandButton cmdAbout 
         Caption         =   "A&bout"
         Height          =   350
         Left            =   120
         TabIndex        =   20
         Top             =   5160
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Enabled         =   0   'False
         Height          =   350
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Upda&te"
         Enabled         =   0   'False
         Height          =   350
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Quit"
         Height          =   350
         Left            =   120
         TabIndex        =   21
         Top             =   5520
         Width           =   1095
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh"
         Height          =   350
         Left            =   120
         TabIndex        =   13
         Top             =   1920
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   350
         Left            =   120
         TabIndex        =   11
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   350
         Left            =   120
         TabIndex        =   12
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   350
         Left            =   120
         TabIndex        =   10
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.TextBox txtFields 
      DataField       =   "BirthDate"
      Height          =   285
      Index           =   4
      Left            =   5160
      MaxLength       =   10
      TabIndex        =   32
      Top             =   1455
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Address"
      Height          =   285
      Index           =   3
      Left            =   1320
      MaxLength       =   50
      TabIndex        =   3
      Top             =   1140
      Width           =   4935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Nippos"
      Height          =   285
      Index           =   2
      Left            =   1320
      MaxLength       =   9
      TabIndex        =   2
      Top             =   825
      Width           =   1215
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Name"
      Height          =   285
      Index           =   1
      Left            =   1320
      MaxLength       =   30
      TabIndex        =   1
      Top             =   495
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "NIM"
      Height          =   285
      Index           =   0
      Left            =   1320
      MaxLength       =   8
      TabIndex        =   0
      Top             =   180
      Width           =   1215
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   36
      Top             =   6240
      Width           =   7995
      _ExtentX        =   14102
      _ExtentY        =   476
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   8625
            MinWidth        =   7408
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "(C) Masino Sinaga (masino_sinaga@yahoo.com)"
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   1764
            MinWidth        =   1764
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "It's up to you..."
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "03/06/2003"
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "Date today"
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            Object.Width           =   1464
            MinWidth        =   1464
            TextSave        =   "16:41"
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "Time right now"
         EndProperty
      EndProperty
   End
   Begin VB.OptionButton optSex 
      Caption         =   "&Male"
      Height          =   255
      Index           =   0
      Left            =   1320
      TabIndex        =   5
      Top             =   1770
      Width           =   975
   End
   Begin VB.OptionButton optSex 
      Caption         =   "&Female"
      Height          =   255
      Index           =   1
      Left            =   2880
      TabIndex        =   6
      Top             =   1770
      Width           =   1215
   End
   Begin VB.Label lblAngka 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   4800
      TabIndex        =   43
      Top             =   5040
      Width           =   1455
   End
   Begin VB.Label lblField 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   120
      TabIndex        =   42
      Top             =   5040
      Width           =   2655
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Marrital Status:"
      Height          =   255
      Index           =   6
      Left            =   240
      TabIndex        =   40
      Top             =   2055
      Width           =   1095
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Sex:"
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   39
      Top             =   1770
      Width           =   855
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Birth Date:"
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   31
      Top             =   1455
      Width           =   855
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Address:"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   30
      Top             =   1140
      Width           =   855
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "ID Employee:"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   29
      Top             =   825
      Width           =   975
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   28
      Top             =   495
      Width           =   855
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "NIM:"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   27
      Top             =   180
      Width           =   855
   End
End
Attribute VB_Name = "frmADOCode2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'File Name  : frmADOCode2.frm
'Description: ADO Code database programming, with
'             add, update, delete, cancel, edit, refresh,
'             navigation (move first, move previous, move
'             next, and move last), find first, find next,
'             filter, sort, and (adjust datagrid's columns
'             width based on the longest field in underlying
'             source) procedure. You can use this code
'             as template for "ADO Code Database
'             Programming" with data validation.
'             Actually, I made this program based on
'             "ADO Code" and "Master Detail" style from
'             "VB Data Form Wizard". I modified by adding
'             Find, Filter, Sort, Bookmark, and
'             Adjust Datagrid's Columns procedure.
'             Reference:
'             - "Microsoft ActiveX Data Objects 2.0 Library"
'             - "Microsoft Data Binding Collection VB 6.0 (SP4)" <--(added automatically by VB6)
'             Component:
'             - "Microsoft Data Grid Control 6.0 (SP5) (OLDEDB)"
'             - "Microsoft Masked Edit Control 6.0 (SP3)"
'             - "Microsoft Windows Common Control 5.0 (SP2)"
'Update     : - The first time I posted this code to
'               www.planet-source-code.com, I still used
'               Indonesia language in its comment. ;)
'             - Now I have translated them to English.
'             - I added a progressbar control to show
'               the progress of adjust datagrid column.
'             - I added criteria "Match whole word only"
'               in checkbox control in Find and Filter
'               procedure.
'Author     : Masino Sinaga (masino_sinaga@yahoo.com)
'Web Site   : http://www30.brinkster.com/masinosinaga/
'             http://www.geocities.com/masino_sinaga/
'Date       : Saturday, 2 November 2002
'Location   : Puslatpos Bandung 40151, INDONESIA
'-------------------------------------------------------

'We use WithEvents because we need them with
'MoveComplete procedure to display record position.
Public WithEvents adoPrimaryRS As Recordset
Attribute adoPrimaryRS.VB_VarHelpID = -1
Public WithEvents rsstrFindData As Recordset
Attribute rsstrFindData.VB_VarHelpID = -1

'General variable...
Dim cekID As Recordset
Attribute cekID.VB_VarHelpID = -1
Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean
Dim strResultBatal As Boolean
Dim NumData As Integer
Dim intRecord As Integer
Dim intField As Integer

Private Sub cboStatus_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{Tab}"
End Sub

Private Sub cmdAbout_Click()
  MsgBox "(c) Masino Sinaga, Bandung - INDONESIA" & vbCrLf & _
         "Saturday, November 2, 2002" & vbCrLf & _
         "E-mail: masino_sinaga@yahoo.com" & vbCrLf & _
         "" & vbCrLf & _
         "I have tried this program before I posted" & vbCrLf & _
         "to www.planet-source-code.com. I didn't" & vbCrLf & _
         "find any error so far. For your information," & vbCrLf & _
         "I use OS WindowsME and VB6SP5." & vbCrLf & _
         "If you find error, please let me know." & vbCrLf & _
         "" & vbCrLf & _
         "You are free to use this code as you want" & vbCrLf & _
         "but please don't forget to write down my" & vbCrLf & _
         "name in your about app if you use my code." & vbCrLf & _
         "Thank you very much. Enjoy!" & vbCrLf & _
         "", vbInformation, "About"
End Sub

Private Sub cmdAbout_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Call Message("About this program.")
End Sub

Private Sub cmdAdd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Call Message("Add new record.")
End Sub

Private Sub Message(strMessage As String)
  StatusBar1.Panels(1).Text = strMessage
End Sub

Private Sub cmdBookmark_Click()
On Error Resume Next
  Set adoBookMark = adoPrimaryRS
  Screen.MousePointer = vbHourglass
  Inisialisasi
  frmBookmark.Show , frmADOCode2
  Screen.MousePointer = vbDefault
End Sub

Private Sub cmdBookmark_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Call Message("Bookmark record so you can go back easily.")
End Sub

Private Sub cmdCancel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Call Message("Cancel the change or new record that have not been saved.")
End Sub

Private Sub cmdClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Call Message("Quit from this program now.")
End Sub

Private Sub cmdDataGrid_Click()
  intRecord = adoPrimaryRS.RecordCount
  intField = adoPrimaryRS.Fields.Count - 1
  Call AdjustDataGridColumnWidth(grdDataGrid, adoPrimaryRS, _
                              intRecord, intField, True)
End Sub

Private Sub cmdDataGrid_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Call Message("Adjust datagrid columns based on the longest field.")
End Sub

Private Sub cmdDelete_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Call Message("Delete the selected record.")
End Sub

Private Sub cmdEdit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Call Message("Edit the selected record.")
End Sub

Private Sub cmdFilter_Click()
On Error Resume Next
  Set adoBookMark = Nothing
  Set adoSort = Nothing
  Set adoFind = Nothing
    
  Set adoFilter = New ADODB.Recordset
  Set adoFilter = adoPrimaryRS
  Screen.MousePointer = vbHourglass
  Inisialisasi
  frmFilter.Show , frmADOCode2
  Screen.MousePointer = vbDefault
End Sub

Private Sub cmdFilter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Call Message("Filter recordset.")
End Sub

Private Sub cmdFind_Click()
On Error Resume Next
  Screen.MousePointer = vbHourglass
  
  Set adoBookMark = Nothing
  Set adoFilter = Nothing
  Set adoSort = Nothing
  
  Set adoFind = New ADODB.Recordset
  Set adoFind = adoPrimaryRS
  Inisialisasi
  frmFind.Show , frmADOCode2
  Screen.MousePointer = vbDefault
End Sub

Private Sub cmdFind_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Call Message("Find record (find first and find next).")
End Sub

Private Sub cmdFirst_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Call Message("Go to the first record.")
End Sub

Private Sub cmdLast_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Call Message("Go to the last record.")
End Sub

Private Sub cmdNext_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Call Message("Go to the next record.")
End Sub

Private Sub cmdPrevious_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Call Message("Go to the previous record.")
End Sub

Private Sub cmdRefresh_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Call Message("Retrieve all records from database.")
End Sub

Private Sub cmdSort_Click()
On Error Resume Next
  Set adoBookMark = Nothing
  Set adoFilter = Nothing
  Set adoFind = Nothing
    
  Set adoSort = New ADODB.Recordset
  Set adoSort = adoPrimaryRS
  Screen.MousePointer = vbHourglass
  Inisialisasi
  frmSort.Show , frmADOCode2
  Screen.MousePointer = vbDefault
End Sub

Private Sub cmdSort_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Call Message("Sort recordset.")
End Sub

Private Sub cmdUnFilter_Click()
  cmdRefresh_Click
  EnablePicture picButtons, True
  cmdUpdate.Enabled = False
  cmdCancel.Enabled = False
End Sub

Private Sub cmdUnFilter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Call Message("Cancel filter recordset = Refresh recordset.")
End Sub

Private Sub cmdUpdate_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Call Message("Save the change or new record.")
End Sub

Private Sub grdDataGrid_Error(ByVal DataError As Integer, Response As Integer)
  Response = -1
  'DataError = -1
End Sub

Private Sub optSex_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = 13 Then cboStatus.SetFocus
End Sub

Public Sub rsstrFindData_MoveComplete(ByVal adReason As _
            ADODB.EventReasonEnum, ByVal pError As _
            ADODB.Error, adStatus As ADODB.EventStatusEnum, _
            ByVal pRecordset As ADODB.Recordset)
    NumData = rsstrFindData.AbsolutePosition
    lblStatus.Caption = "Record number " & CStr(NumData) & " " & _
                        "of " & rsstrFindData.RecordCount
End Sub

Private Sub Form_Load()
On Error GoTo Message
  INIFileName = App.Path & "\SettingADOCode2.ini"
  strResultBatal = False
  UnlockTheFormKoneksi
  Set adoPrimaryRS = New Recordset
  'We display all data in a datagrid below and underlying
  'source (the selected record in datagrid) above.
  adoPrimaryRS.Open "SHAPE {select NIM,Name,Nippos,Address," & _
                    "BirthDate,Sex,Status from t_mhs Order by NIM} " & _
                    "AS ParentCMD APPEND ({select NIM," & _
                    "Name,Nippos,Address,BirthDate,Sex,Status FROM t_mhs " & _
                    "ORDER BY NIM } AS ChildCMD RELATE NIM " & _
                    "TO NIM) AS ChildCMD", db, _
                    adOpenDynamic, _
                    adLockOptimistic
  Dim oText As TextBox
  'Bind textbox to recordset
  For Each oText In Me.txtFields
    Set oText.DataSource = adoPrimaryRS
  Next
  'Bind recordset to datagrid
  Set grdDataGrid.DataSource = adoPrimaryRS.DataSource
  mbDataChanged = False
  LockTheForm  'Lock textbox, and make datagrid enable
  grdDataGrid.Enabled = True
  'If we have no data in recordset
  If adoPrimaryRS.RecordCount < 1 Then
     MsgBox "Recordset is empty!", vbExclamation, _
            "Empty Recordset"
     UnlockTheForm     'Unlock textbox so we can add new record
     cmdAdd_Click
  End If
  'Fill in cboStatus here
  cboStatus.AddItem "Single"
  cboStatus.AddItem "Marry"
  cboStatus.AddItem "Widower"
  cboStatus.AddItem "Widow"
  
  'Display description to user in the form, but we still
  'keep the code from sex and marrital status data
  SexDescription
  MarriageStatus
  GetDate
  
  LockTheForm 'Lock textbox, combobox, and optionbutton
  'Except Datagrid....
  grdDataGrid.Enabled = True
  
  'Make sure you put following statement in this procedure
  'in order that to prevent error if we navigate
  'recordset by clicking DataGrid and move pointer
  'through DataGrid or Navigation button.
  'I put following statement both here and property window
  'just to make sure I won't forget this.
  'If I don't put this one, I get an error duplicate
  'key in primary key (field) in DataGrid, because
  'DataGrid will make a duplicate data in primary field
  'when I move recordset by clicking cmdFirst or
  'cmdLast button.
  grdDataGrid.TabStop = False
  
  Exit Sub
Message:
  MsgBox Err.Number & " - " & Err.Description
  End
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If cmdUpdate.Enabled = True And cmdCancel.Enabled = True Then
     MsgBox "You have to save or cancel the changes " & vbCrLf & _
            "that you have just made before quit!", _
            vbExclamation, "Warning"
     cmdUpdate.SetFocus
     Cancel = -1
     Exit Sub
  End If
  'Clear all memory from object variable
  If Not adoFind Is Nothing Then
     Set adoFind = Nothing
  ElseIf Not adoFilter Is Nothing Then
     Set adoFilter = Nothing
  ElseIf Not adoSort Is Nothing Then
     Set adoSort = Nothing
  ElseIf Not adoBookMark Is Nothing Then
    Set adoBookMark = Nothing
  End If
  If Not adoPrimaryRS Is Nothing Then _
    Set adoPrimaryRS = Nothing  'Clear memory from recordset
  'In order that prevent error from DataGrid...!
  If grdDataGrid.TabStop = True Then
     txtFields(0).SetFocus
  End If
  db.Close 'Close database
  Set db = Nothing  'Clear memory from database
  End
End Sub

'If user hit the button in keyboard, so get the code
'from Form_KeyDown
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If mbEditFlag Or mbAddNewFlag Then Exit Sub
  Select Case KeyCode
    Case vbKeyEscape
      cmdClose_Click
    Case vbKeyEnd
      cmdLast_Click
    Case vbKeyHome
      cmdFirst_Click
    Case vbKeyUp, vbKeyPageUp
      If Shift = vbCtrlMask Then
        cmdFirst_Click
      Else
        cmdPrevious_Click
      End If
    Case vbKeyDown, vbKeyPageDown
      If Shift = vbCtrlMask Then
        cmdLast_Click
      Else
        cmdNext_Click
      End If
  End Select
End Sub

'Make mouse pointer back to normal
Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

'Display the selected record in datagrid
Public Sub adoPrimaryRS_MoveComplete(ByVal adReason As _
            ADODB.EventReasonEnum, ByVal pError As _
            ADODB.Error, adStatus As ADODB.EventStatusEnum, _
            ByVal pRecordset As ADODB.Recordset)
  NumData = adoPrimaryRS.AbsolutePosition
  lblStatus.Caption = "Record number " & CStr(NumData) & " from " _
                      & adoPrimaryRS.RecordCount
  CheckNavigation
End Sub

Private Sub CheckNavigation()
  'This will check which navigation button can be
  'accessed when you navigate the recordset through
  'Datagrid control or navigation button itself
  With adoPrimaryRS
   'If we have at least two record...
   If (.RecordCount > 1) Then
      'BOF = Begin Of Recordset
      If (.BOF) Or _
         (.AbsolutePosition = 1) Then
          cmdFirst.Enabled = False
          cmdPrevious.Enabled = False
          cmdNext.Enabled = True
          cmdLast.Enabled = True
      'EOF = End Of Recordset
      ElseIf (.EOF) Or _
          (.AbsolutePosition = .RecordCount) Then
          cmdNext.Enabled = False
          cmdLast.Enabled = False
          cmdFirst.Enabled = True
          cmdPrevious.Enabled = True
      Else
          cmdFirst.Enabled = True
          cmdPrevious.Enabled = True
          cmdNext.Enabled = True
          cmdLast.Enabled = True
      End If
   Else
      cmdFirst.Enabled = False
      cmdPrevious.Enabled = False
      cmdNext.Enabled = False
      cmdLast.Enabled = False
   End If
 End With
End Sub

Public Sub adoPrimaryRS_WillChangeRecord(ByVal adReason As _
            ADODB.EventReasonEnum, ByVal cRecords As Long, _
            adStatus As ADODB.EventStatusEnum, _
            ByVal pRecordset As ADODB.Recordset)
  'Besides in each procedure, you can make validation here
  'This is the event raised when following happen
  
  Dim bCancel As Boolean
  Select Case adReason
  Case adRsnAddNew
  Case adRsnClose
  Case adRsnDelete
  Case adRsnFirstChange
  Case adRsnMove
  Case adRsnRequery
  Case adRsnResynch
  Case adRsnUndoAddNew
  Case adRsnUndoDelete
  Case adRsnUndoUpdate
  Case adRsnUpdate
  End Select
  If bCancel Then adStatus = adStatusCancel
End Sub

Private Sub cmdAdd_Click()
  On Error GoTo AddErr
  With adoPrimaryRS
    If Not (.BOF And .EOF) Then
      mvBookMark = .Bookmark
    End If
    UnlockTheForm
    .AddNew
    lblStatus.Caption = "Add record"
    mbAddNewFlag = True
    SetButtons False
  End With
  'Default value alway the first (top)
  optSex(0).Value = True
  cboStatus.Text = cboStatus.List(0)
  EmptyBirthDate
  grdDataGrid.Enabled = False  'In order that prevent error
  On Error Resume Next
  txtFields(0).SetFocus
  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmdDelete_Click()
  On Error GoTo DeleteErr
  If MsgBox("Are you sure you want to delete this record?", _
            vbQuestion + vbYesNo + vbDefaultButton2, _
            "Delete Record") _
            <> vbYes Then
     Exit Sub
  End If
  With adoPrimaryRS
    .Delete
    .MoveNext
    If .EOF Then .MoveLast
  End With
  SexDescription
  MarriageStatus
  Exit Sub
DeleteErr:
  MsgBox Err.Description
End Sub

Private Sub cmdRefresh_Click()
  'Refresh is very important in multiuser app
  On Error GoTo RefreshErr
  Set adoFind = Nothing
  Set adoFilter = Nothing
  Set adoSort = Nothing
  Set adoBookMark = Nothing
  If strResultBatal = True Then
     SetButtons True
     strResultBatal = False
  End If
  SexDescription
  MarriageStatus
  LockTheForm
  cmdBookmark.Enabled = True
  cmdAdd.Enabled = True
  cmdEdit.Enabled = True
  cmdDelete.Enabled = True
  cmdRefresh.Enabled = True
  Set grdDataGrid.DataSource = Nothing
  Set adoPrimaryRS = New Recordset
  adoPrimaryRS.Open "SHAPE {select NIM,Name,Nippos,Address," & _
                    "BirthDate,Sex,Status from t_mhs Order by NIM} " & _
                    "AS ParentCMD APPEND ({select NIM," & _
                    "Name,Nippos,Address,BirthDate,Sex,Status FROM t_mhs " & _
                    "ORDER BY NIM } AS ChildCMD RELATE NIM " & _
                    "TO NIM) AS ChildCMD", db, _
                    adOpenStatic, adLockOptimistic
  
  Dim oText As TextBox
  'Bind textbox to recordset
  For Each oText In Me.txtFields
    Set oText.DataSource = adoPrimaryRS
  Next
  Set grdDataGrid.DataSource = adoPrimaryRS.DataSource
  grdDataGrid.Enabled = True
  Exit Sub
RefreshErr:
  cmdAdd.Enabled = False
  cmdEdit.Enabled = False
  cmdUpdate.Enabled = False
  cmdDelete.Enabled = False
  cmdCancel.Enabled = False
  cmdRefresh.Enabled = True
  mbEditFlag = False
  mbAddNewFlag = False
  adoPrimaryRS.CancelUpdate
  If mvBookMark <> 0 Then
      adoPrimaryRS.Bookmark = mvBookMark
  Else
      adoPrimaryRS.MoveFirst
  End If
  cmdBookmark.Enabled = True
  mbDataChanged = False
  strResultBatal = True
  cmdRefresh_Click  'Automatically refresh
  Exit Sub
End Sub

Private Sub cmdEdit_Click()
  On Error GoTo EditErr
  lblStatus.Caption = "Edit record"
  mbEditFlag = True
  SetButtons False
  UnlockTheForm 'Unlock textbox in order that we can edit data
  txtFields(0).SetFocus: SendKeys "{Home}+{End}"
  Exit Sub
EditErr:
  MsgBox Err.Description
End Sub

Private Sub cmdCancel_Click()
  On Error Resume Next
  LockTheForm
  cmdRefresh_Click
  grdDataGrid.Enabled = True
  If strResultBatal = True Then
     Exit Sub
  End If
  SetButtons True
  mbEditFlag = False
  mbAddNewFlag = False
  adoPrimaryRS.CancelUpdate
  If mvBookMark > 0 Then
    adoPrimaryRS.Bookmark = mvBookMark
  Else
    adoPrimaryRS.MoveFirst
  End If
  SexDescription
  MarriageStatus
  GetDate
  LockTheForm    'Lock textbox
  grdDataGrid.Enabled = True
  mbDataChanged = False
End Sub

Private Sub cmdUpdate_Click()
Dim i As Integer
  On Error GoTo UpdateErr
  For i = 0 To 3
    If txtFields(i).Text = "" Then
       MsgBox "You have to fill in all textbox!", _
              vbExclamation, "Validation"
       txtFields(i).SetFocus
       Exit Sub
    End If
  Next i
  ReceiveBirthDate  'Get birth date value from entry form
  If Not IsDate(mskTglLahir.Text) Then
     MsgBox "Birth date or its format is invalid!", _
            vbExclamation, "Birth Date"
     mskTglLahir.SetFocus: SendKeys "{Home}+{End}"
     Exit Sub
  End If
  If optSex(0).Value = True And _
     cboStatus.Text = "Widow" Then
     MsgBox "There is no Widow with sex Male..." & vbCrLf & _
            "Change the sex or Marrital Status!", _
            vbExclamation, "Invalid"
     cboStatus.SetFocus
     Exit Sub
  End If
  If optSex(1).Value = True And _
     cboStatus.Text = "Widower" Then
     MsgBox "There is no Widower with sex Female..." & vbCrLf & _
            "Change the sex or Marrital Status!", _
            vbExclamation, "Invalid"
     cboStatus.SetFocus
     Exit Sub
  End If
  'Check double data in primary key (field)
  Set cekID = New Recordset
  cekID.Open "SELECT * FROM t_mhs WHERE NIM=" & _
             "'" & Trim(txtFields(0).Text) & "'", db
  If cekID.RecordCount > 0 And mbAddNewFlag Then
     MsgBox "NIM '" & txtFields(0).Text & "' already exist. " & vbCrLf & _
            "Please change to another NIM!", _
            vbExclamation, "Double NIM"
     txtFields(0).SetFocus: SendKeys "{Home}+{End}"
     Set cekID = Nothing
     Exit Sub
  End If
  'Retrieve the code from Sex and Marrital Status
  AssignSex
  AssignStatus
  'Update by using UpdateBatch. UpdateBatch will
  'automatically update all data in various fields type.
  adoPrimaryRS.UpdateBatch adAffectAll
  'Move pointer to last record if we just added data
  If mbAddNewFlag Then
    adoPrimaryRS.MoveLast
  End If
  
  'Modified in Tuesday, February 18, 2003
  'based on information from Asep Jaelani
  'through e-mail to me.
  'If we edit data, Birth Date can be saved
  'without moving pointer in recordset manually
  If mbEditFlag Then
    adoPrimaryRS.MoveNext
    adoPrimaryRS.MovePrevious
    GetDate
  End If
  '---------------------------------------------
  'After all, get description of sex and marriage
  SexDescription
  MarriageStatus
  'Update all status
  mbEditFlag = False
  mbAddNewFlag = False
  SetButtons True
  mbDataChanged = False
  LockTheForm  'Lock textbox
  grdDataGrid.Enabled = True
  'Display the record position
  NumData = adoPrimaryRS.AbsolutePosition
  lblStatus.Caption = "Record number " & CStr(NumData) & _
                      " of " & adoPrimaryRS.RecordCount
  Exit Sub
UpdateErr:
  Select Case Err.Number
         Case -2147467259
              MsgBox "NIM '" & txtFields(0).Text & "' already exist." & vbCrLf & _
                     "Please change to another NIM!", _
                    vbExclamation, "Double NIM"
              txtFields(0).SetFocus: SendKeys "{Home}+{End}"
         Case Else
              MsgBox Err.Number & " - " & _
                     Err.Description, vbCritical, "Error"
  End Select
End Sub

Private Sub mskTglLahir_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{Tab}"
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdFirst_Click()
  On Error GoTo GoFirstError
  If adoFilter Is Nothing Then
     adoPrimaryRS.MoveFirst
  Else
     adoFilter.MoveFirst
  End If
  mbDataChanged = False
  SexDescription
  MarriageStatus
  GetDate
  Exit Sub
GoFirstError:
  MsgBox Err.Description
End Sub

Private Sub cmdLast_Click()
  On Error GoTo GoLastError
  If adoFilter Is Nothing Then
     adoPrimaryRS.MoveLast
  Else
     adoFilter.MoveLast
  End If
  mbDataChanged = False
  SexDescription
  MarriageStatus
  GetDate
  Exit Sub
GoLastError:
  MsgBox Err.Description
End Sub

Private Sub cmdNext_Click()
  On Error GoTo GoNextError
  If adoFilter Is Nothing Then
     If Not adoPrimaryRS.EOF Then adoPrimaryRS.MoveNext
     If adoPrimaryRS.EOF And adoPrimaryRS.RecordCount > 0 Then
        Beep
        adoPrimaryRS.MoveLast
        MsgBox "This is the last record.", _
               vbInformation, "Last Record"
     End If
  Else
     If Not adoFilter.EOF Then adoFilter.MoveNext
     If adoFilter.EOF And adoFilter.RecordCount > 0 Then
        Beep
        adoFilter.MoveLast
        MsgBox "This is the last record.", _
               vbInformation, "Last Record"
     End If
  End If
  mbDataChanged = False
  SexDescription
  MarriageStatus
  GetDate
  Exit Sub
GoNextError:
  MsgBox Err.Description
End Sub

Private Sub cmdPrevious_Click()
  On Error GoTo GoPrevError
  If adoFilter Is Nothing Then
     If Not adoPrimaryRS.BOF Then adoPrimaryRS.MovePrevious
     If adoPrimaryRS.BOF And adoPrimaryRS.RecordCount > 0 Then
        Beep
        adoPrimaryRS.MoveFirst
        MsgBox "This is the first record.", _
               vbInformation, "First Record"
     End If
  Else
     If Not adoFilter.BOF Then adoFilter.MovePrevious
     If adoFilter.BOF And adoFilter.RecordCount > 0 Then
        Beep
        adoFilter.MoveFirst
        MsgBox "This is the first record.", _
               vbInformation, "First Record"
     End If
  End If
  mbDataChanged = False
  SexDescription
  MarriageStatus
  GetDate
  Exit Sub
GoPrevError:
  MsgBox Err.Description
End Sub

Private Sub SetButtons(bVal As Boolean)
  'Adjust which buttons will be activated or be locked
  cmdAdd.Enabled = bVal
  cmdEdit.Enabled = bVal
  cmdUpdate.Enabled = Not bVal
  cmdCancel.Enabled = Not bVal
  cmdDelete.Enabled = bVal
  cmdClose.Enabled = bVal
  cmdRefresh.Enabled = bVal
  cmdNext.Enabled = bVal
  cmdFirst.Enabled = bVal
  cmdLast.Enabled = bVal
  cmdPrevious.Enabled = bVal
  cmdFilter.Enabled = bVal
  cmdUnFilter.Enabled = bVal
  cmdSort.Enabled = bVal
  cmdFind.Enabled = bVal
  cmdBookmark.Enabled = bVal
  cmdDataGrid.Enabled = bVal
  cmdAbout.Enabled = bVal
  cmdClose.Enabled = bVal
End Sub

'If we click DataGrid, then adjust all field above with
'the selected record ini recordset in DataGrid
Private Sub grdDataGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error Resume Next 'To prevent error from GetDate proc.
  'If mbAddNewFlag = False (cmdAdd.Enabled = True)
  'then get sex description and marriage status
  If cmdAdd.Enabled = True Then
     SexDescription
     MarriageStatus
     GetDate
  Else 'If mbAddNewFlag = True (cmdAdd.Enabled = False)
       'then just make maskedbox empty. We can not
       'get the sex description and marriage status
       'because it will raise an error...
     EmptyBirthDate  'Just make maskedbox empty here
  End If
End Sub

Private Sub txtFields_KeyPress(Index As Integer, KeyAscii As Integer)
  Select Case Index  'If we hit Enter, jump to next textbox
         Case 0 To 3
              If KeyAscii = 13 Then SendKeys "{Tab}"
  End Select
End Sub

'Get the sex description based on the code from
'database, then change it to optionbutton
Private Sub SexDescription()
  If txtFields(5).Text = "L" Then
     optSex(0).Value = True
  Else
     optSex(1).Value = True
  End If
End Sub

'This is reverse of SexDescription, we get the code
'based on the optionbutton value
Private Sub AssignSex()
  If optSex(0).Value = True Then
     txtFields(5).Text = "L"
  Else
     txtFields(5).Text = "P"
  End If
End Sub

'Get the marriage description based on the code from
'database, then change it to optionbutton
Private Sub MarriageStatus()
  If txtFields(6).Text = "B" Then
     cboStatus.Text = "Single"
  ElseIf txtFields(6).Text = "M" Then
     cboStatus.Text = "Marry"
  ElseIf txtFields(6).Text = "D" Then
     cboStatus.Text = "Widower"
  ElseIf txtFields(6).Text = "J" Then
     cboStatus.Text = "Widow"
  End If
End Sub

'This is reverse of MarriageStatus, we get the code
'based on the optionbutton value
Private Sub AssignStatus()
  If cboStatus.Text = "Single" Then
     txtFields(6).Text = "B"
  ElseIf cboStatus.Text = "Marry" Then
     txtFields(6).Text = "M"
  ElseIf cboStatus.Text = "Widower" Then
     txtFields(6).Text = "D"
  ElseIf cboStatus.Text = "Widow" Then
     txtFields(6).Text = "J"
  End If
End Sub

'Lock textbox in order that we can't edit data
Private Sub LockTheForm()
Dim i As Integer
  For i = 0 To 6
    txtFields(i).Locked = True
  Next i
  mskTglLahir.Enabled = False
  optSex(0).Enabled = False
  optSex(1).Enabled = False
  cboStatus.Locked = True
  grdDataGrid.Enabled = False
End Sub

'Unlock textbox in order that we can edit data
Sub UnlockTheForm()
Dim i As Integer
  For i = 0 To 6
    txtFields(i).Locked = False
  Next i
  mskTglLahir.Enabled = True
  optSex(0).Enabled = True
  optSex(1).Enabled = True
  cboStatus.Locked = False
  grdDataGrid.Enabled = False
End Sub

'Get birth date from database, then take it to maskedbox
'in order that we can display it to maskedbox format
Private Sub GetDate()
  'Updated at June 3, 2003, by adding Format function
  'if datetime setting in user's computer not use
  'dd/mm/yyyy. (Example: d/m/yyyy or d/m/yy or dd/mm/yy, etc)
  'In this example, I use Mask property in MaskEdBox
  '= ##/##/#### or dd/mm/yyyy (setting in Indonesia).
  'If you want to change this property value,
  'please change the "dd/mm/yyyy" to your own setting
  'for example: "mm/dd/yyyy" in Regional Setting
  'at Date tab...
  mskTglLahir.Text = Format(txtFields(4).Text, "dd/mm/yyyy")
  'This will caused an error "Invalid property value"
  'because if date setting in regional setting does
  'not use dd/mm/yyyy or mm/dd/yyyy, then the property
  'value of maskedbox is not match with maskedbox value
  'mskTglLahir.Text = txtFields(4).Text
End Sub

'This is reverse of GetDate, in order that we can save
'birth date to database with maskedbox format
Private Sub ReceiveBirthDate()
  txtFields(4).Text = mskTglLahir.Text
End Sub

'Make maskedbox empty, but we have to assign text property
'with maskedbox format or mask like this:
Private Sub EmptyBirthDate()
  mskTglLahir.Text = "__/__/____"
End Sub
