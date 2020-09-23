VERSION 5.00
Begin VB.Form frmBookmark 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bookmark"
   ClientHeight    =   4215
   ClientLeft      =   8175
   ClientTop       =   3600
   ClientWidth     =   3210
   Icon            =   "frmBookmark.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   3210
   Begin VB.CommandButton cmdButton 
      Caption         =   "&Help"
      Height          =   375
      Index           =   4
      Left            =   1080
      TabIndex        =   7
      ToolTipText     =   "What is bookmark? How do I use it?"
      Top             =   3720
      Width           =   975
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   3
      Left            =   2160
      TabIndex        =   6
      ToolTipText     =   "Finish with bookmark"
      Top             =   3720
      Width           =   975
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "&Jump"
      Height          =   375
      Index           =   2
      Left            =   2160
      TabIndex        =   5
      ToolTipText     =   "Go to record which bookmark name selected"
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "&Delete"
      Height          =   375
      Index           =   1
      Left            =   1080
      TabIndex        =   4
      ToolTipText     =   "Delete the selected bookmark"
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "&Add"
      Height          =   375
      Index           =   0
      Left            =   50
      TabIndex        =   3
      ToolTipText     =   "Add new bookmark"
      Top             =   3240
      Width           =   975
   End
   Begin VB.ListBox lstBookmark 
      Height          =   2400
      Left            =   50
      TabIndex        =   2
      ToolTipText     =   "Double click the name to go to its record"
      Top             =   720
      Width           =   3135
   End
   Begin VB.TextBox txtBookmark 
      Height          =   285
      Left            =   50
      MaxLength       =   30
      TabIndex        =   0
      ToolTipText     =   "Enter bookmark name here"
      Top             =   360
      Width           =   3135
   End
   Begin VB.Label Label1 
      Caption         =   "Bookmark name:"
      Height          =   255
      Left            =   50
      TabIndex        =   1
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "frmBookmark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'File Name  : frmBookmark.frm
'Description: Mark a record in recordset so we can
'             go back to that record later without
'             remember the position of record.
'Author     : Masino Sinaga (masino_sinaga@yahoo.com)
'Web Site   : http://www30.brinkster.com/masinosinaga/
'             http://www.geocities.com/masino_sinaga/
'Date       : Sunday, November 3, 2002
'Location   : Puslatpos Bandung 40151, INDONESIA
'-----------------------------------------------------

'User Defined Type (UDT) for bookmark
Private Type arrMark
   AbsolutePosition As Double
   BookmarkName As String * 30
   BookmarkNumber As Variant
End Type

'Declare dynamic array with arrMark types
Dim tabMark() As arrMark

'This is procedure to give mark to a record
Private Sub GiveMark()
On Error GoTo Message
'Static, so we can increase this variable as long as
'program stay in memory even we declare them in procedure
Static intNumber As Integer
'Add to array each time user add a new bookmark
ReDim Preserve tabMark(UBound(tabMark) + 1)
'Update counter variable intNumber
intNumber = intNumber + 1
  'Get information for this bookmark we added
  tabMark(intNumber).AbsolutePosition = adoBookMark.AbsolutePosition
  tabMark(intNumber).BookmarkNumber = adoBookMark.Bookmark
  tabMark(intNumber).BookmarkName = txtBookmark.Text
  Exit Sub
Message:
  MsgBox Err.Number & " - " & Err.Description
End Sub

Private Sub cmdButton_Click(Index As Integer)
  Select Case Index
         Case 0  'Add button clicked
              Dim i As Integer
              For i = 0 To lstBookmark.ListCount - 1
                If lstBookmark.List(i) = txtBookmark.Text Then
                   MsgBox "This bookmark name already exist in the list!" & vbCrLf & _
                          "Please change to another name...", _
                          vbExclamation, "Bookmark Name"
                   txtBookmark.SetFocus
                   SendKeys "{Home}+{End}"
                   Exit Sub
                 End If
              Next i
              lstBookmark.AddItem txtBookmark.Text
              GiveMark
              txtBookmark.Text = ""
              cmdButton(0).Enabled = False
         Case 1 'Delete button clicked
              If lstBookmark.ListCount > 0 Then
                 lstBookmark.RemoveItem lstBookmark.ListIndex
                 If lstBookmark.ListCount > 0 Then
                    lstBookmark.Selected(0) = True
                    cmdButton(1).Enabled = True
                 Else
                    cmdButton(1).Enabled = False
                    cmdButton(2).Enabled = False
                 End If
              Else
                 cmdButton(1).Enabled = False
                 cmdButton(2).Enabled = False
              End If
         Case 2 'Jump button clicked
              Dim strTemp As String
              Dim Location As Double
              strTemp = Trim(lstBookmark.List(lstBookmark.ListIndex))
              Location = CekintPosition(strTemp)
              'Here is the essential of bookmark we added,
              'we can jump direct to position we bookmark
              'before...
              adoBookMark.MoveFirst
              adoBookMark.Move Location - 1
         Case 3 'Cancel button clicked
              Me.Hide 'Just hide this form in order that
                      'we can use it later as long as program
                      'stay in memory
         Case 4 'Help button clicked, display how to use
                'bookmark...
              MsgBox "1. Bookmak is a way to mark a record in a recordset " & vbCrLf & _
                     "   so you can go back to the record later quickly" & vbCrLf & _
                     "   without remember the position of that record." & vbCrLf & _
                     "   Let program keep the position of record." & vbCrLf & _
                     "" & vbCrLf & _
                     "2. Select the record you want to bookmark by" & vbCrLf & _
                     "   clicking it in DataGrid or through Navigation " & vbCrLf & _
                     "   button in frmADOCode2, then enter the bookmark " & vbCrLf & _
                     "   name in the textbox above, and press Enter or " & vbCrLf & _
                     "   click 'Add' button to add this name to the listbox " & vbCrLf & _
                     "   below. This will keep/save your bookmark." & vbCrLf & _
                     "" & vbCrLf & _
                     "3. If you want to go back to record that you have" & vbCrLf & _
                     "   marked, click bookmark name in the listbox" & vbCrLf & _
                     "   then click 'Jump' button, or you can double-click" & vbCrLf & _
                     "   the bookmark name in the listbox. " & vbCrLf & _
                     "" & vbCrLf & _
                     "4. If you want to delete the bookmark name, click" & vbCrLf & _
                     "   bookmark name in the listbox, then click" & vbCrLf & _
                     "   'Delete' button." & vbCrLf & _
                     "" & vbCrLf & _
                     "", vbInformation, "About Bookmark and How To Use It"
                     
  End Select
End Sub

'This will check and take the position of bookmark
Function CekintPosition(Name As String) As Double
Dim i As Integer
  For i = 0 To UBound(tabMark)
    If Name = Trim(tabMark(i).BookmarkName) Then
       CekintPosition = tabMark(i).AbsolutePosition
       Exit For
    End If
  Next i
End Function

'In order that there is no double bookmark name saved in the listbox
Private Sub CheckDouble()
Dim i As Integer
  For i = 0 To lstBookmark.ListCount - 1
    If lstBookmark.List(i) = txtBookmark.Text Then
       MsgBox "This bookmark name already exist in the list!" & vbCrLf & _
              "You can not save the same bookmark name." & vbCrLf & _
              "Please change to another name...", _
              vbExclamation, "Bookmark Name"
              txtBookmark.SetFocus
       SendKeys "{Home}+{End}"
       Exit Sub
    End If
  Next i

End Sub

Private Sub Form_Load()
  LockTheFormButton
  ReDim tabMark(lstBookmark.ListCount)
  'Get setting for this form from INI File
  Call ReadFromINIToControls(frmBookmark, "Bookmark")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Cancel = -1  'We don't unload this form
  Me.Hide      'Just hide them in order that we can
               'use this form later as long as program
               'stays in memory
End Sub

'If a bookmark name was selected or clicked...
Private Sub lstBookmark_Click()
  If lstBookmark.ListCount > 0 Then
     'Unlock the button we will use
     UnlockTheFormButton
  End If
  'If textbox is empty, Add button is not active
  'Add button is not active in order that there is
  'no bookmark name contains a empty string
  If Len(Trim(txtBookmark.Text)) = 0 Then _
     cmdButton(0).Enabled = False
End Sub

'Alternative way to go to the record we mark,
'by double click the bookmark name in listbox
Private Sub lstBookmark_DblClick()
  cmdButton(2).Enabled = True
  cmdButton_Click (2)
End Sub

'If there is a change in textbox
Private Sub txtBookmark_Change()
  'If textbox is not empty
  If Len(Trim(txtBookmark.Text)) > 0 Then
     'Add button is active and ready now
     cmdButton(0).Enabled = True
     cmdButton(0).Default = True
  Else 'If textbox is empty
     cmdButton(0).Enabled = False
  End If
End Sub

'Lock the button that we don't use it
Private Sub LockTheFormButton()
  For i = 0 To 2
    cmdButton(i).Enabled = False
  Next i
End Sub

'Unlock the button that we need
Private Sub UnlockTheFormButton()
  For i = 0 To 2
    cmdButton(i).Enabled = True
  Next i
End Sub
