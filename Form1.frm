VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10875
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   10875
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView ListView1 
      Height          =   5175
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   9128
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.OptionButton SOrder 
      Caption         =   "Descending"
      Height          =   285
      Index           =   1
      Left            =   2610
      TabIndex        =   2
      Top             =   5340
      Width           =   1395
   End
   Begin VB.OptionButton SOrder 
      Caption         =   "Ascending"
      Height          =   285
      Index           =   0
      Left            =   1350
      TabIndex        =   1
      Top             =   5340
      Width           =   1245
   End
   Begin VB.Label lblSort 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   5910
      TabIndex        =   4
      Top             =   5370
      Width           =   4785
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Sort Order:"
      Height          =   255
      Left            =   30
      TabIndex        =   3
      Top             =   5370
      Width           =   1245
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Sub Form_Load()
  Dim oListItem As ListItem
  Dim dblDate As Double
  ListView1.View = lvwReport
  ListView1.Sorted = True

  With ListView1.ColumnHeaders
    .Add Text:="AlphaNumeric"
    .Add Text:="Date"
    .Add Text:="Numeric 1"
    .Add Text:="Numeric 2"
    .item(1).Width = 2000
    .item(2).Width = 2000
    .item(3).Width = 2000
    .item(4).Width = 2000
    .item(2).Alignment = lvwColumnRight
    .item(3).Alignment = lvwColumnRight
    .item(4).Alignment = lvwColumnRight
  End With
 
  With ListView1.ListItems
    Do While .Count < 1000
      If .Count < 500 Then
        Set oListItem = .Add(, , "Item: (" & .Count + 1 & ")")
      ElseIf .Count < 750 Then
        Set oListItem = .Add(, , "ABC (" & .Count - 499 & ")")
      Else
        Set oListItem = .Add(, , "123" & .Count + 1)
      End If
      dblDate = Rnd * 10
      oListItem.ListSubItems.Add , , CDate(dblDate)
      dblDate = Rnd * 1000000000
      oListItem.ListSubItems.Add , , CDbl(dblDate)
      dblDate = Rnd * 100000000
      oListItem.ListSubItems.Add , , Format(dblDate, "########")
      dblDate = Rnd * 10
      oListItem.ListSubItems.Add , , "Item: " & .Count * 5
    Loop
  End With
  SOrder(0).Value = True
End Sub
Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  Dim tStart As Long
  tStart = GetTickCount
  
  ListView1.Visible = False
  ListView1.SortKey = ColumnHeader.Index - 1
  
  If SOrder(0).Value Then
    ListView1.SortOrder = lvwAscending
  Else
    ListView1.SortOrder = lvwDescending
  End If

  Select Case ColumnHeader.Index
    Case 1: SortListView ListView1, LVNatural
    Case 2: SortListView ListView1, LVDate
    Case 3, 4: SortListView ListView1, LVNumeric
  End Select
  
  ListView1.Visible = True
  lblSort.Caption = "Sorted in: " & GetTickCount - tStart & " ms"
End Sub
