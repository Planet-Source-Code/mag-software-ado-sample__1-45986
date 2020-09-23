VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmCustomers 
   BackColor       =   &H8000000A&
   Caption         =   "www.wiz-solutions.com"
   ClientHeight    =   5190
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8715
   Icon            =   "frmCustomers.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5190
   ScaleWidth      =   8715
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton pbFind 
      Caption         =   "&Find"
      Height          =   375
      Left            =   6180
      TabIndex        =   29
      Top             =   3000
      Width           =   1155
   End
   Begin VB.CommandButton pbEnd 
      Cancel          =   -1  'True
      Caption         =   "&Exit"
      Height          =   525
      Left            =   7080
      TabIndex        =   19
      Top             =   3420
      Width           =   1245
   End
   Begin VB.CommandButton pbEdit 
      Caption         =   "&Edit"
      Height          =   375
      Left            =   3870
      TabIndex        =   18
      Top             =   3000
      Width           =   1155
   End
   Begin VB.CommandButton pbDelete 
      Caption         =   "&Delete"
      Height          =   375
      Left            =   5040
      TabIndex        =   17
      Top             =   3000
      Width           =   1155
   End
   Begin VB.CommandButton pbSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   1335
      TabIndex        =   16
      Top             =   3000
      Width           =   1155
   End
   Begin VB.CommandButton pbCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2490
      TabIndex        =   15
      Top             =   3000
      Width           =   1155
   End
   Begin VB.CommandButton pbAddNew 
      Caption         =   "Add &new"
      Height          =   375
      Left            =   180
      TabIndex        =   14
      Top             =   3000
      Width           =   1155
   End
   Begin VB.CommandButton pbMoveNext 
      Caption         =   ">"
      Height          =   345
      Left            =   1710
      TabIndex        =   13
      Top             =   3660
      Width           =   765
   End
   Begin VB.CommandButton pbMovePrev 
      Caption         =   "<"
      Height          =   345
      Left            =   945
      TabIndex        =   12
      Top             =   3660
      Width           =   765
   End
   Begin VB.CommandButton pbMoveLast 
      Caption         =   ">>"
      Height          =   345
      Left            =   2475
      TabIndex        =   11
      Top             =   3660
      Width           =   765
   End
   Begin VB.CommandButton pbMoveFirst 
      Caption         =   "<<"
      Height          =   345
      Left            =   180
      TabIndex        =   10
      Top             =   3660
      Width           =   765
   End
   Begin VB.TextBox txtFax 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5520
      TabIndex        =   6
      Tag             =   "2"
      Text            =   "Text1"
      Top             =   1380
      Width           =   1635
   End
   Begin VB.TextBox txtCompanyName 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1470
      TabIndex        =   1
      Tag             =   "1"
      Text            =   "Text1"
      Top             =   690
      Width           =   4425
   End
   Begin VB.TextBox txtContactName 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   390
      TabIndex        =   2
      Tag             =   "2"
      Text            =   "Text1"
      Top             =   1380
      Width           =   3315
   End
   Begin VB.TextBox txtAddress 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   390
      MultiLine       =   -1  'True
      TabIndex        =   4
      Tag             =   "2"
      Text            =   "frmCustomers.frx":014A
      Top             =   2100
      Width           =   3375
   End
   Begin VB.TextBox txtCity 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6240
      TabIndex        =   7
      Tag             =   "2"
      Text            =   "Text1"
      Top             =   2070
      Width           =   1635
   End
   Begin VB.TextBox txtState 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3840
      TabIndex        =   8
      Tag             =   "2"
      Text            =   "Text1"
      Top             =   2070
      Width           =   675
   End
   Begin VB.TextBox txtPostalCode 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4560
      TabIndex        =   3
      Tag             =   "2"
      Text            =   "Text1"
      Top             =   2070
      Width           =   1635
   End
   Begin VB.TextBox txtPhone 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3810
      TabIndex        =   5
      Tag             =   "2"
      Text            =   "Text1"
      Top             =   1380
      Width           =   1635
   End
   Begin MSAdodcLib.Adodc adoPublishers 
      Height          =   330
      Left            =   4470
      Top             =   3660
      Visible         =   0   'False
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Customers"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox txtCompanyID 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   390
      TabIndex        =   0
      Tag             =   "0"
      Text            =   "Text1"
      Top             =   690
      Width           =   885
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fax:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   11
      Left            =   5520
      TabIndex        =   28
      Top             =   1170
      Width           =   345
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Phone:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   10
      Left            =   3840
      TabIndex        =   27
      Top             =   1170
      Width           =   570
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Postal code:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   8
      Left            =   4560
      TabIndex        =   26
      Top             =   1860
      Width           =   1020
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Region:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   7
      Left            =   3870
      TabIndex        =   25
      Top             =   1860
      Width           =   630
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "City:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   6
      Left            =   6240
      TabIndex        =   24
      Top             =   1860
      Width           =   375
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   420
      TabIndex        =   23
      Top             =   1890
      Width           =   735
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contact name:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   420
      TabIndex        =   22
      Top             =   1140
      Width           =   1230
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Company name:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   1500
      TabIndex        =   21
      Top             =   480
      Width           =   1365
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Company ID"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   330
      TabIndex        =   20
      Top             =   480
      Width           =   1035
   End
   Begin VB.Label lblPosition 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6420
      TabIndex        =   9
      Top             =   60
      Width           =   2025
   End
End
Attribute VB_Name = "frmCustomers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------Module header sample----------------------------------
'
' Form: frmCustomers
' File: frmCustomers.frm
' Author: dUGI
' E.mail: mmilak@net4u.hr
' Web: www.wiz-solutions.com
' Dates:
' 29.05.2003 - Created

' Purpose:
'     how to use ado data control
'------------------------------------------------------------------------
Option Explicit

'recordset edit status
Private Enum EditMode
   NowNone = 0    'browsing
   NowEdit = 1    'editing record
   NowAdd = 2     'add record
   NowFind = 3    'find record
End Enum

'txtBox coloring
Private Enum EditColors
   NotAllowed = 0 'user entry not allowed (Autonumber fields...)
   Required = 1   'required fields (Company name...)
   Custom = 2     'free entry (at will...)
End Enum

Dim CurrentEditStatus      As EditMode       '
Dim FieldType              As EditColors     '

Private Sub Form_Load()

'setting data control
With adoPublishers
   Debug.Print App.Path
   .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source = " & App.Path & "\Biblio2000.mdb"
   .RecordSource = "Publishers"
   
   'activate data control & fill recordset with data
   .Refresh
End With




'set status --> record browsing
ChangeStatus NowNone

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

If CurrentEditStatus <> NowNone And CurrentEditStatus <> NowFind Then
   MsgBox "Cannot exit while editing or adding a record!!!", vbCritical
   Cancel = True
End If

End Sub

Private Sub pbAddNew_Click()

adoPublishers.Recordset.AddNew
ChangeStatus NowAdd
txtCompanyID.SetFocus

End Sub

Private Sub pbCancel_Click()
adoPublishers.Recordset.CancelUpdate
ChangeStatus NowNone
End Sub

Private Sub pbDelete_Click()
Dim iodgovor As Byte

iodgovor = MsgBox("Really delete current record?", vbQuestion + vbYesNo)
If iodgovor = vbYes Then
   With adoPublishers.Recordset
      .Delete
      
      'change current record
      If .RecordCount > 0 Then      'is there any record?
         If .BOF Then               'is it first record deleted?
            .MoveFirst
         Else
            .MovePrevious
         End If
      Else
         'Record count = 0
         UpdateButtons
      End If
      
   End With
End If
End Sub

Private Sub pbEdit_Click()
ChangeStatus NowEdit
txtCompanyName.SetFocus
End Sub

Private Sub pbEnd_Click()
Unload Me
End Sub

Private Sub pbFind_Click()
Dim iPosition As Integer
Dim sName As String

sName = InputBox("Enter company name")
iPosition = adoPublishers.Recordset.AbsolutePosition

If sName = "" Then Exit Sub

ChangeStatus NowFind
adoPublishers.Recordset.Find "[Company name] like '" & sName & "*'"

If adoPublishers.Recordset.EOF Then
   MsgBox "No match"
   adoPublishers.Recordset.AbsolutePosition = iPosition
End If
ChangeStatus NowNone

End Sub

Private Sub pbMoveFirst_Click()
adoPublishers.Recordset.MoveFirst
ChangeStatus NowNone
End Sub

Private Sub pbMoveLast_Click()
adoPublishers.Recordset.MoveLast
ChangeStatus NowNone
End Sub

Private Sub pbMoveNext_Click()
adoPublishers.Recordset.MoveNext
ChangeStatus NowNone
End Sub

Private Sub pbMovePrev_Click()
adoPublishers.Recordset.MovePrevious
ChangeStatus NowNone
End Sub

Private Sub pbSave_Click()

'data validation check
If fnValidateInput Then
   SaveData
      
   'update status
   ChangeStatus NowNone
End If

End Sub

Private Function fnValidateInput() As Boolean
'data validation check

'assume ok
fnValidateInput = True

'test company name
If Len(txtCompanyName.Text) = 0 Then
   MsgBox "Incorrect Company name entry", vbCritical
   txtCompanyName.SetFocus
   fnValidateInput = False
   Exit Function
End If

End Function

Private Sub ChangeStatus(NewState As EditMode)

CurrentEditStatus = NewState

Select Case NewState
   Case NowAdd
      ReadData
      lblPosition.Caption = "Adding record"
      
   Case NowEdit
      ReadData
      lblPosition.Caption = "Editing record"
      
   Case NowNone
      ReadData
   
   Case NowFind
      lblPosition.Caption = "Finding record"
      
   Case Else
      MsgBox "????"
      
End Select

UpdateButtons
UpdateTextBoxes

End Sub

Private Sub UpdateButtons()

Select Case CurrentEditStatus
   Case NowAdd, NowEdit
      pbAddNew.Enabled = False
      pbEdit.Enabled = False
      pbDelete.Enabled = False
      pbSave.Enabled = True
      pbCancel.Enabled = True
      pbFind.Enabled = False
      
      'navigation buttons
      pbMoveFirst.Enabled = False
      pbMoveLast.Enabled = False
      pbMoveNext.Enabled = False
      pbMovePrev.Enabled = False
   
   Case NowNone
      pbAddNew.Enabled = True
      pbEdit.Enabled = True
      pbDelete.Enabled = True
      pbSave.Enabled = False
      pbCancel.Enabled = False
      pbFind.Enabled = True
   
      With adoPublishers.Recordset
         'current record test
         If .AbsolutePosition = 1 Then
            pbMoveFirst.Enabled = False
            pbMovePrev.Enabled = False
            pbMoveNext.Enabled = True
            pbMoveLast.Enabled = True
         ElseIf .AbsolutePosition = .RecordCount Then
            pbMoveFirst.Enabled = True
            pbMovePrev.Enabled = True
            pbMoveNext.Enabled = False
            pbMoveLast.Enabled = False
         Else
            pbMoveFirst.Enabled = True
            pbMovePrev.Enabled = True
            pbMoveNext.Enabled = True
            pbMoveLast.Enabled = True
         End If
      End With
   
End Select
End Sub

Private Sub UpdateTextBoxes()
Dim ctl As Control

For Each ctl In Me.Controls
   If TypeOf ctl Is TextBox Then
      If CurrentEditStatus = NowNone Or CurrentEditStatus = NowFind Then
         ctl.Enabled = False
         ctl.BackColor = vbWhite
         Unload frmLegend
      Else
         ctl.Enabled = True
         'colors legend
         frmLegend.Show
         
         'depending on tag property, change control color
         Select Case ctl.Tag
            Case NotAllowed
               ctl.BackColor = CLR_RED
               ctl.Locked = True
            Case Required
               ctl.BackColor = CLR_YELLOW
               
            Case Custom
               ctl.BackColor = CLR_GREEN
               
         End Select
      End If
   End If
Next ctl

End Sub

Private Sub ReadData()
'read from recordset

With adoPublishers.Recordset
   lblPosition.Caption = "Record " & .AbsolutePosition & " of " & .RecordCount
   
   txtCompanyID.Text = !PubID & vbNullString
   txtCompanyName.Text = ![Company name] & vbNullString
   txtContactName.Text = !Name & vbNullString
   txtFax.Text = !Fax & vbNullString
   txtPhone.Text = !Telephone & vbNullString
   txtPostalCode.Text = !Zip & vbNullString
   txtState.Text = !State & vbNullString
   txtCity.Text = !City & vbNullString
   txtAddress.Text = !Address & vbNullString
   
End With

End Sub



Private Sub SaveData()
'read data from textboxes and save

'data validation
 If fnValidateInput = True Then
   'read from textboxes
   With adoPublishers.Recordset
      ![Company name] = txtCompanyName.Text
      !Name = txtContactName.Text
      !Address = txtAddress.Text
      !State = txtState.Text
      !Fax = txtFax.Text
      !Telephone = txtPhone.Text
      !Zip = txtPostalCode.Text
      
      'update method
      adoPublishers.Recordset.Update
      
   End With
   'message
   MsgBox "Record saved :)", vbExclamation
End If

End Sub
