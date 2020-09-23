VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form Form1 
   Caption         =   "Search"
   ClientHeight    =   8460
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11745
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8460
   ScaleWidth      =   11745
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Search For"
      Height          =   645
      Left            =   3990
      TabIndex        =   7
      Top             =   60
      Width           =   7695
      Begin VB.CommandButton Command3 
         Caption         =   "Change View"
         Height          =   255
         Left            =   6120
         TabIndex        =   13
         Top             =   240
         Width           =   1125
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4440
         TabIndex        =   11
         Text            =   "*.*"
         Top             =   300
         Width           =   915
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Stop"
         Height          =   285
         Left            =   3660
         TabIndex        =   10
         Top             =   240
         Width           =   555
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   9
         Text            =   ".bmp"
         Top             =   240
         Width           =   2745
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Start"
         Height          =   285
         Left            =   2970
         TabIndex        =   8
         Top             =   240
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Filter"
         Height          =   195
         Left            =   4440
         TabIndex        =   12
         Top             =   120
         Width           =   330
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Search Results"
      Height          =   7635
      Left            =   4020
      TabIndex        =   5
      Top             =   690
      Width           =   7665
      Begin MSComctlLib.ListView ListView1 
         Height          =   7305
         Left            =   90
         TabIndex        =   6
         Top             =   240
         Width           =   7425
         _ExtentX        =   13097
         _ExtentY        =   12885
         SortKey         =   1
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "File Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "File Path"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Set Start Directory And Watch Search In Action"
      Height          =   8265
      Left            =   30
      TabIndex        =   0
      Top             =   60
      Width           =   3885
      Begin VB.DirListBox Dir1 
         Appearance      =   0  'Flat
         Height          =   3015
         Left            =   90
         TabIndex        =   4
         Top             =   720
         Width           =   3675
      End
      Begin VB.FileListBox File1 
         Appearance      =   0  'Flat
         Height          =   3150
         Hidden          =   -1  'True
         Left            =   90
         System          =   -1  'True
         TabIndex        =   3
         Top             =   3780
         Width           =   3675
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   1005
         Left            =   90
         TabIndex        =   2
         Top             =   6990
         Width           =   3675
      End
      Begin VB.DriveListBox Drive1 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   90
         TabIndex        =   1
         Top             =   360
         Width           =   3705
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3630
      Top             =   4530
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":6852
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":76A4
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
On Error GoTo bere
Dim num As Long
Me.Tag = "Start"

'== RESET PAST RESULTS =======
ListView1.ListItems.Clear
List1.Clear
'=============================



'== ADD CURRENT DIRECTORY TO DIRECOTRY TO SEARCH =====
List1.AddItem Dir1.Path
List1.ListIndex = 0
'==============================================


ere:
If Me.Tag = "Stop" Then GoTo bere

Me.Caption = "Search " & ListView1.ListItems.Count & " Results"

'== GO TO THE NEXT PATH TO SEARCH
Dir1.Path = List1.Text
'===========================

'== ADD SUB FOLDERS TO LIST OF FOLDERS TO SEARCH ===
'== EXTRACT FOLDER NAME FROM FOLDER PATH & SEARCH IN THE FOLDER NAME FOR SEARCH STRING
    'IF SEARCH STRING FOUND ADDED TO LISTVIEW
For a = 0 To Dir1.ListCount - 1
List1.AddItem Dir1.List(a)
If InStr(1, UCase(GetLast(Dir1.List(a))), UCase(Text1.Text)) > 0 Then
Set lv = ListView1.ListItems.Add(, , GetLast(Dir1.List(a)), 1, 1)
lv.ListSubItems.Add , , Dir1.Path
End If
DoEvents
Next a
'============================================================

'== SERACH ALL FILES FOR SERACH STRING IF FOUND ADD TO LISTVIEW
For i = 0 To File1.ListCount - 1
If InStr(1, UCase(File1.List(i)), UCase(Text1.Text)) > 0 Then
Set lv = ListView1.ListItems.Add(, , File1.List(i), 2, 2)
lv.ListSubItems.Add , , Dir1.Path
End If
DoEvents
Next i
'=====================================

'== TO SET NEXT PATH  TO CHANGE TO
List1.ListIndex = List1.ListIndex + 1
GoTo ere
'================================


'==SEARCH IS NOW COMPLETE
bere:
MsgBox "COMPLETE"
'==============================




End Sub

Private Sub Command2_Click()
Me.Tag = "Stop"
End Sub

Private Sub Command3_Click()
If ListView1.View = 0 Then
ListView1.View = 3
Else
On Error Resume Next
ListView1.View = ListView1.View - 1
End If
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
Dir1.Refresh
File1.Refresh
End Sub

Private Sub Text2_Change()
On Error Resume Next
File1.Pattern = Text2.Text
End Sub


Function GetLast(dir As String) As String
txt = Split(dir, "\")
GetLast = txt(UBound(txt))
End Function

Private Sub Drive1_Change()
On Error Resume Next
Dir1.Path = Drive1.Drive
Dir1.Refresh
File1.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub
