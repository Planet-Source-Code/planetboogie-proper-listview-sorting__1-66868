Attribute VB_Name = "modLVSort"
Option Explicit
Option Compare Text
Public Enum LVSortEnum
 LVNatural = 0
 LVNumeric = 1
 LVDate = 2
End Enum
Private LV As ListView
Private LVSort As LVSortEnum
Private Const SORT_DESCENDING = &H80000000
Private Const SORT_COLUMNMASK = &HFF
Private Const LVM_FIRST = &H1000
Private Const LVM_SORTITEMS = (LVM_FIRST + 48)
Private Const WM_DESTROY = &H2
Private Const LVProc = (-4)
Private Const OLDLVProc = "OldLVProc"
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)
Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Sub SortListView(ByVal LView As ListView, ByVal SortType As LVSortEnum)
 LVSort = SortType
 Set LV = LView
 With LV
  .Sorted = False
  Call LVSubClass(.hWnd, AddressOf ListViewProc)
  .Sorted = True
  Call UnLVSubClass(.hWnd)
 End With
 Set LV = Nothing
End Sub
Private Static Function LVCompare(ByVal lParam1 As Long, ByVal lParam2 As Long, ByVal lParamSort As Long) As Long
 Dim LVCol As Long
 Dim RetVal As Long
 Dim Val1 As String
 Dim Val2 As String
 LVCol = lParamSort And SORT_COLUMNMASK
 Select Case LVCol
  Case 0
   Val1 = GetLVItem(lParam1).Text
   Val2 = GetLVItem(lParam2).Text
  Case Else
   Val1 = GetLVItem(lParam1).SubItems(LVCol)
   Val2 = GetLVItem(lParam2).SubItems(LVCol)
 End Select
 Select Case LVSort
  Case LVNatural: RetVal = StrCompFileNames(Val1, Val2)
  Case LVNumeric: RetVal = CCur(Val1) - CCur(Val2)
  Case LVDate:    RetVal = CDate(Val1) - CDate(Val2)
 End Select
 Select Case CBool(lParamSort And SORT_DESCENDING)
  Case True: LVCompare = -RetVal
  Case False: LVCompare = RetVal
 End Select
End Function
Private Function GetLVItem(lParam As Long) As ListItem
 Dim lpli As Long
 Dim li As ListItem
 If lParam Then
  Call MoveMemory(lpli, ByVal lParam + 8, 4)
  If lpli Then
   Call MoveMemory(li, lpli, 4)
   Set GetLVItem = li
   Call MoveMemory(li, 0&, 4)
  End If
 End If
End Function
Private Function LVSubClass(hWnd As Long, lpfnNew As Long) As Boolean
 Dim lpfnOld As Long
 Dim fSuccess As Boolean
 If GetProp(hWnd, OLDLVProc) Then
  LVSubClass = True
  Exit Function
 Else
  lpfnOld = SetWindowLong(hWnd, LVProc, lpfnNew)
  If lpfnOld Then LVSubClass = SetProp(hWnd, OLDLVProc, lpfnOld)
 End If
End Function
Private Function UnLVSubClass(hWnd As Long) As Boolean
 Dim lpfnOld As Long
 lpfnOld = GetProp(hWnd, OLDLVProc)
 If lpfnOld Then
  If RemoveProp(hWnd, OLDLVProc) Then
   UnLVSubClass = SetWindowLong(hWnd, LVProc, lpfnOld)
  End If
 End If
End Function
Private Function ListViewProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
 Select Case uMsg
  Case LVM_SORTITEMS
   wParam = LV.SortKey Or (CBool(LV.SortOrder) And SORT_DESCENDING)
   lParam = FARPROC(AddressOf LVCompare)
  Case WM_DESTROY
   Call CallWindowProc(GetProp(hWnd, OLDLVProc), hWnd, uMsg, wParam, lParam)
   Call UnLVSubClass(hWnd)
   Exit Function
 End Select
 ListViewProc = CallWindowProc(GetProp(hWnd, OLDLVProc), hWnd, uMsg, wParam, lParam)
End Function
Private Function FARPROC(pfn As Long) As Long
 FARPROC = pfn
End Function
