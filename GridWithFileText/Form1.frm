VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4935
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7575
   LinkTopic       =   "Form1"
   ScaleHeight     =   4935
   ScaleWidth      =   7575
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   4560
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "3/14/02"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   7575
      TabIndex        =   10
      Top             =   0
      Width           =   7575
      Begin VB.CommandButton cmdSave 
         Caption         =   "&SaveToFile"
         Height          =   375
         Left            =   6360
         TabIndex        =   8
         ToolTipText     =   "Save To File"
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   5040
         TabIndex        =   7
         ToolTipText     =   "Delete"
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         Height          =   375
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Update"
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&AddNew"
         Height          =   375
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Add New"
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdLast 
         Caption         =   ">>"
         Height          =   375
         Left            =   1560
         TabIndex        =   4
         ToolTipText     =   "Last"
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   ">"
         Height          =   375
         Left            =   1080
         TabIndex        =   3
         ToolTipText     =   "Next"
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "<"
         Height          =   375
         Left            =   600
         TabIndex        =   2
         ToolTipText     =   "Previous"
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdFirst 
         Caption         =   "<<"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         ToolTipText     =   "First"
         Top             =   120
         Width           =   495
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   873
      _Version        =   393216
      AllowArrows     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      TabAcrossSplits =   -1  'True
      TabAction       =   2
      WrapCellPointer =   -1  'True
      AllowDelete     =   -1  'True
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
            LCID            =   1033
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
            LCID            =   1033
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
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuSave 
         Caption         =   "&SaveToFile"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
         Shortcut        =   ^E
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'The purpose of the following code is to demonstrate
'how you can load a datagrid with data from a text file and
'write them back to the file.
'Project > References > "Microsoft ActiveX data object 2.0 library"
'Project > References>"Microsoft Scripting Runtime"


Option Explicit
Dim rs As Recordset
Const conWrite As String = 2
Const conRead As String = 1
Private Sub cmdAdd_Click()
'disable the command buttons
ButEnable False
rs.MoveLast
rs.AddNew
DataGrid1.SetFocus
End Sub

Private Sub cmdDelete_Click()
rs.Delete
rs.MoveNext
If rs.EOF Then
    rs.MoveLast
End If
' count the records
NumRec
End Sub

Private Sub cmdFirst_Click()
rs.MoveFirst
End Sub

Private Sub cmdLast_Click()
rs.MoveLast
End Sub

Private Sub cmdNext_Click()
rs.MoveNext
If rs.EOF Then
  rs.MoveLast
End If

End Sub

Private Sub cmdPrevious_Click()
rs.MovePrevious
If rs.BOF Then
    rs.MoveFirst
End If
End Sub

Private Sub cmdSave_Click()
SaveToFile
End Sub

Private Sub cmdUpdate_Click()
rs.Update
'enable the command buttons
ButEnable True
'count the records
NumRec
End Sub

Private Sub Form_Load()
'create rs and bind grid to it
If CreateRs(OpenFile(conRead)) Then
Set DataGrid1.DataSource = rs
'count the records
NumRec
Else
MsgBox "Application cannot create a recordset."
End If

End Sub

Private Function CreateRs(ts As TextStream) As Boolean
Dim s
On Error GoTo errs
'get reference to rs
Set rs = New Recordset
'build fields
With rs
    .Fields.Append "ID", adBSTR
    .Fields.Append "LastName", adBSTR
    .Fields.Append "FirstName", adBSTR
    .Fields.Append "Phone", adBSTR
    .Fields.Append "Address", adBSTR
    .Fields.Append "City", adBSTR
    .Fields.Append "State", adBSTR
    .Fields.Append "Zip", adBSTR
    .Fields.Append "Contract", adBSTR
    .Open
End With
'dump data to rs
Dim I As Integer
Do While ts.AtEndOfStream <> True
   s = Split(ts.ReadLine, vbTab)
  With rs
    .AddNew
    For I = 0 To 8
    .Fields(I) = s(I)
    Next
    .Update
   End With
Loop
'we are happy
CreateRs = True
Exit Function
errs:
'failure

CreateRs = False
End Function

Private Sub SaveToFile()
Dim l As String
  
rs.MoveFirst
'convert rs to a string
l = rs.GetString(adClipString, rs.RecordCount, vbTab, vbCrLf, "Null")
' write the string to the file
OpenFile(conWrite).Write (l)
End Sub
Private Function OpenFile(x As String) As TextStream
Dim fso As New FileSystemObject
Dim fil As File
Dim ts As TextStream
Dim s As String
Dim strPath As String
'get a path
strPath = App.Path & "\Grid.txt"
'open  the file
    Set fil = fso.GetFile(strPath)
    Set ts = fil.OpenAsTextStream(x)
    Set OpenFile = ts
End Function

Private Sub Form_Resize()
'Resize datgrid1.
If Form1.ScaleHeight <> 0 Then
DataGrid1.Width = Form1.ScaleWidth - 150

DataGrid1.Height = Form1.ScaleHeight / 1.5
End If
End Sub
Private Sub ButEnable(v As Boolean)
Dim obj As Object

For Each obj In Me.Controls
 If TypeOf obj Is CommandButton Then
   If obj.Caption <> "&Update" Then
    obj.Enabled = v
    End If
End If
Next
End Sub
Private Sub NumRec()
StatusBar1.Panels(1).Text = "Number Of Records:  " & rs.RecordCount
End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub mnuSave_Click()
SaveToFile
End Sub







