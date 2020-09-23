VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSimpleArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'//************************************
'//     Author: Morgan Haueisen
'//             morganh@hartcom.net
'//************************************

'// Speed Test: 1.5 GHz, 512MB of RAM, running XP Pro
'// Collection vs Class (non-sorted)
'// Adding 80,000 Records:
'//    Collection:    114 ms
'//         Class:    297 ms
'// Reading 80,000 Records
'//    Collection: 172318 ms
'//         Class:    190 ms
'//
'// VB's ListBox control can only hold 32767 items
'//        Adding: 2157 ms
'//       Reading:  286 ms

Private mstrItemArray() As String
Private mlngDataArray() As Long

Private mlngNewIndex    As Long
Private mlngListCount   As Long
Private mblnSorted      As Boolean
Private Const C_STEP    As Long = 100 '// uses a little more memory but is faster then adding 1 at a time
Private Const C_SERROR  As String = "The selected index number exceeds the number of elements dimensioned in the array!"

Public Sub AddItem(ByVal vstrItem As String, _
                   Optional ByVal vlngIndex As Long = -1)

  Dim lngI As Long

   On Error GoTo Err_Proc

   '// Does the index requested exceed the size of the array?
   If Not (vlngIndex < 0) Then
      If vlngIndex > mlngListCount + 1 Then
         Err_Handler True, 0, C_SERROR, "clsSimpleArray", "AddItem"
         GoTo Exit_Proc
      ElseIf vlngIndex = mlngListCount + 1 Then
         vlngIndex = -1
      End If
   End If
   
   '// Resize the array
   If mlngListCount < 0 Then '// Does the array exist?
      ReDim mstrItemArray(0 To C_STEP) As String
      ReDim mlngDataArray(0 To C_STEP) As Long
      mlngListCount = 0
      mlngNewIndex = 0
   
   Else
      If mblnSorted Then '// Find insert point if sorted
         For lngI = 0 To mlngListCount
            If vstrItem < mstrItemArray(lngI) Then
               vlngIndex = lngI
               Exit For
            End If
         Next lngI
      End If

      mlngListCount = mlngListCount + 1
      mlngNewIndex = mlngListCount
      If mlngListCount > UBound(mstrItemArray) Then
         ReDim Preserve mstrItemArray(0 To mlngListCount + C_STEP) As String
         ReDim Preserve mlngDataArray(0 To mlngListCount + C_STEP) As Long
      End If
   End If

   '// Insert new item
   If Not (vlngIndex < 0) Then
      '// Shift data right to make a space to insert new data
      For lngI = mlngListCount To vlngIndex + 1 Step -1
         mstrItemArray(lngI) = mstrItemArray(lngI - 1)
         mlngDataArray(lngI) = mlngDataArray(lngI - 1)
      Next lngI
      mlngNewIndex = vlngIndex
   End If
   
   '// Add data to array
   mstrItemArray(mlngNewIndex) = vstrItem
   mlngDataArray(mlngNewIndex) = 0
   
Exit_Proc:

   Exit Sub

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "clsSimpleArray", "AddItem"
   Err.Clear
   Resume Exit_Proc

End Sub

Private Sub Class_Initialize()
   
   mlngListCount = -1
   mlngNewIndex = -1

End Sub

Private Sub Class_Terminate()
   
   Erase mstrItemArray
   Erase mlngDataArray

End Sub

Public Sub Clear()

   Erase mstrItemArray
   Erase mlngDataArray
   mlngListCount = -1
   mlngNewIndex = -1

End Sub

Private Sub Err_Handler(Optional ByVal vblnDisplayError As Boolean = True, _
                        Optional ByVal vstrErrNumber As String = vbNullString, _
                        Optional ByVal vstrErrDescription As String = vbNullString, _
                        Optional ByVal vstrModuleName As String = vbNullString, _
                        Optional ByVal vstrProcName As String = vbNullString)

  Dim strTemp As String
  Dim lngFN   As Long

   '// Purpose: Error handling - On Error
   On Error Resume Next

   '// Show Error Message
   If vblnDisplayError Then
      strTemp = "Error occured: "
      If LenB(vstrErrNumber) Then strTemp = strTemp & vstrErrNumber & vbNewLine Else strTemp = strTemp & vbNewLine
      If LenB(vstrErrDescription) Then strTemp = strTemp & "Description: " & vstrErrDescription & vbNewLine
      If LenB(vstrModuleName) Then strTemp = strTemp & "Module: " & vstrModuleName & vbNewLine
      If LenB(vstrProcName) Then strTemp = strTemp & "Function: " & vstrProcName
      MsgBox strTemp, vbCritical, App.Title & " - ERROR"
   End If

   '// Write error log
   lngFN = FreeFile
   Open App.Path & "\ErrorLog.txt" For Append As #lngFN
   Write #lngFN, Now, vstrErrNumber, vstrErrDescription, vstrModuleName, vstrProcName, _
      App.Title & " v" & App.Major & "." & App.Minor & "." & App.Revision, _
      Environ$("username"), Environ$("computername")
   Close #lngFN

End Sub

Public Function FindItem(ByVal vstrItem As String, _
                         Optional ByVal vblnExactMatch As Boolean = False) As Boolean

  Dim lngI As Long

   '// Sets NewIndex if successful
   
   On Error Resume Next
   mlngNewIndex = -1

   If Not (mlngListCount < 0) Then '// Has the array been created?

      For lngI = 0 To mlngListCount

         If vblnExactMatch Then
            '// Look for Exact match
            If UCase$(mstrItemArray(lngI)) = UCase$(vstrItem) Then
               FindItem = True
               mlngNewIndex = lngI
               Exit For
            End If

         Else
            '// Look for Like match
            If InStrB(1, UCase$(mstrItemArray(lngI)), UCase$(vstrItem)) Then
               FindItem = True
               mlngNewIndex = lngI
               Exit For
            End If
         End If

      Next lngI

   End If

End Function

Public Function FindItemData(ByVal vlngItemData As Long) As Boolean

  Dim lngI As Long

   '// Sets NewIndex if successful
   
   On Error Resume Next
   mlngNewIndex = -1

   If Not (mlngListCount < 0) Then '// Has the array been created?

      For lngI = 0 To mlngListCount
         If mlngDataArray(lngI) = vlngItemData Then
            FindItemData = True
            mlngNewIndex = lngI
            Exit For
         End If
      Next lngI

   End If

End Function

Public Property Get ItemData(ByVal vlngIndex As Long) As Long

   If vlngIndex < 0 Or mlngListCount < vlngIndex Then
      Err_Handler True, 0, C_SERROR, "clsSimpleArray", "Get ItemData"
   Else
      ItemData = mlngDataArray(vlngIndex)
   End If

End Property

Public Property Let ItemData(ByVal vlngIndex As Long, ByVal vNewValue As Long)

   If vlngIndex < 0 Or mlngListCount < vlngIndex Then
      Err_Handler True, 0, C_SERROR, "clsSimpleArray", "Let ItemData"
   Else
      mlngDataArray(vlngIndex) = vNewValue
   End If

End Property

Public Property Get List(ByVal vlngIndex As Long) As String

   On Error Resume Next
   If vlngIndex < 0 Or mlngListCount < vlngIndex Then
      Err_Handler True, 0, C_SERROR, "clsSimpleArray", "Get AddItem"
   Else
      List = mstrItemArray(vlngIndex)
   End If

End Property

Public Property Let List(ByVal vlngIndex As Long, vNewValue As String)

   On Error Resume Next
   If vlngIndex < 0 Or mlngListCount < vlngIndex Then
      Err_Handler True, 0, C_SERROR, "clsSimpleArray", "Let AddItem"
   Else
      mstrItemArray(vlngIndex) = vNewValue
   End If
   
End Property

Public Property Get ListCount() As Long

   If mlngListCount < 0 Then
      ListCount = 0
   Else
      ListCount = mlngListCount + 1
   End If

End Property

Public Property Get NewIndex() As Long

   NewIndex = mlngNewIndex

End Property

Public Sub RemoveDuplicates()

  Dim lngI As Long
  Dim lngN As Long

   On Error Resume Next

   If mlngListCount > 0 Then '// more then 1 record?

      For lngI = 0 To mlngListCount
         If lngI > mlngListCount Then Exit For
         
         For lngN = 0 To mlngListCount

            If lngN > mlngListCount Then Exit For
            If Not (lngN = lngI) Then
               '// Remove item if it appears more than once.
               If UCase$(mstrItemArray(lngN)) = UCase$(mstrItemArray(lngI)) Then
                  Call RemoveItem(lngN)
                  lngN = lngN - 1
               End If
            End If

         Next lngN
      Next lngI
      
   End If

End Sub

Public Sub RemoveItem(ByVal vlngIndex As Long)

  Dim lngI As Long

   On Error GoTo Err_Proc

   If vlngIndex < 0 Or vlngIndex > mlngListCount Then
      '// The index requested exceeds the size of the array.
      Err_Handler True, 0, C_SERROR, "clsSimpleArray", "RemoveItem"
      
   Else
      If mlngListCount > 0 Then
         '// shift data left
         For lngI = vlngIndex To mlngListCount
            If lngI + 1 <= mlngListCount Then
               mstrItemArray(lngI) = mstrItemArray(lngI + 1)
               mlngDataArray(lngI) = mlngDataArray(lngI + 1)
            End If
         Next lngI
      
         mstrItemArray(mlngListCount) = vbNullString
         mlngDataArray(mlngListCount) = 0
      End If
      
      mlngListCount = mlngListCount - 1
      mlngNewIndex = mlngListCount
      
      If mlngListCount > 0 Then
         If mlngListCount + C_STEP < UBound(mstrItemArray) Then
            ReDim Preserve mstrItemArray(0 To mlngListCount) As String
            ReDim Preserve mlngDataArray(0 To mlngListCount) As Long
         End If
         
      Else
         Erase mstrItemArray
         Erase mlngDataArray
      End If
   
   End If

Exit_Proc:
   Exit Sub

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "clsSimpleArray", "RemoveItem"
   Err.Clear
   Resume Exit_Proc

End Sub

Public Sub RemoveItemsMatching(ByVal vstrItem As String)

  Dim lngI As Long

   On Error Resume Next

   If Not (mlngListCount < 0) Then '// Has the array been created?
      
      For lngI = 0 To mlngListCount
         If lngI > mlngListCount Then Exit For
         If UCase$(mstrItemArray(lngI)) = UCase$(vstrItem) Then
            Call RemoveItem(lngI)
            lngI = lngI - 1
         End If
      Next lngI
   
   End If
   
End Sub

Public Sub RemoveItemDataMatching(ByVal vlngItemData As Long)

  Dim lngI As Long

   On Error Resume Next

   If Not (mlngListCount < 0) Then '// Has the array been created?
   
      For lngI = 0 To mlngListCount
         If lngI > mlngListCount Then Exit For
         If mlngDataArray(lngI) = vlngItemData Then
            Call RemoveItem(lngI)
            lngI = lngI - 1
         End If
      Next lngI
   
   End If
   
End Sub

Public Property Let Sorted(ByVal vblnNewValue As Boolean)

   mblnSorted = vblnNewValue

End Property

Public Property Get Sorted() As Boolean

   Sorted = mblnSorted

End Property

Public Property Get Text() As String
   
   If Not (mlngNewIndex < 0) Then
      Text = mstrItemArray(mlngNewIndex)
   End If

End Property

