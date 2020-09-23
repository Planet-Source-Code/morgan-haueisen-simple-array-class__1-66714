VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6885
   ClientLeft      =   3210
   ClientTop       =   2145
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6885
   ScaleWidth      =   6585
   Begin VB.CheckBox chkSorted 
      Caption         =   "Sorted"
      Height          =   225
      Left            =   4905
      TabIndex        =   20
      Top             =   3075
      Width           =   1275
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Fill List1"
      Height          =   525
      Left            =   4875
      TabIndex        =   19
      Top             =   3435
      Width           =   1275
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Speed ListBox"
      Height          =   525
      Left            =   4050
      TabIndex        =   13
      Top             =   1470
      Width           =   1860
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Speed Collection"
      Height          =   525
      Left            =   4050
      TabIndex        =   8
      Top             =   810
      Width           =   1860
   End
   Begin VB.ListBox List2 
      Height          =   2985
      Left            =   2835
      TabIndex        =   4
      Top             =   2955
      Width           =   1800
   End
   Begin VB.CommandButton Command2 
      Caption         =   "RemoveDup"
      Height          =   540
      Left            =   4905
      TabIndex        =   3
      Top             =   4185
      Width           =   1245
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Speed SimpleArray"
      Height          =   525
      Left            =   4050
      TabIndex        =   2
      Top             =   165
      Width           =   1860
   End
   Begin VB.ListBox List1 
      Height          =   2985
      Left            =   765
      TabIndex        =   1
      Top             =   2955
      Width           =   1800
   End
   Begin VB.Label Label5 
      Caption         =   "Using SimpleArray to do the work and ListBox to display the results"
      Height          =   1230
      Left            =   4845
      TabIndex        =   21
      Top             =   4935
      Width           =   1380
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Speed test based on 32767 records"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1890
      TabIndex        =   18
      Top             =   2235
      Width           =   4320
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   ":Read"
      Height          =   195
      Index           =   2
      Left            =   3495
      TabIndex        =   17
      Top             =   1770
      Width           =   435
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   ":Add"
      Height          =   195
      Index           =   2
      Left            =   3495
      TabIndex        =   16
      Top             =   1530
      Width           =   330
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Height          =   195
      Index           =   5
      Left            =   3375
      TabIndex        =   15
      Top             =   1770
      Width           =   45
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Height          =   195
      Index           =   4
      Left            =   3375
      TabIndex        =   14
      Top             =   1530
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   ":Read"
      Height          =   195
      Index           =   1
      Left            =   3495
      TabIndex        =   12
      Top             =   1110
      Width           =   435
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   ":Add"
      Height          =   195
      Index           =   1
      Left            =   3495
      TabIndex        =   11
      Top             =   870
      Width           =   330
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Height          =   195
      Index           =   3
      Left            =   3375
      TabIndex        =   10
      Top             =   1110
      Width           =   45
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Height          =   195
      Index           =   2
      Left            =   3375
      TabIndex        =   9
      Top             =   870
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   ":Read"
      Height          =   195
      Index           =   0
      Left            =   3495
      TabIndex        =   7
      Top             =   465
      Width           =   435
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   ":Add"
      Height          =   195
      Index           =   0
      Left            =   3495
      TabIndex        =   6
      Top             =   225
      Width           =   330
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Height          =   195
      Index           =   1
      Left            =   3375
      TabIndex        =   5
      Top             =   465
      Width           =   45
   End
   Begin VB.Line Line1 
      X1              =   60
      X2              =   6360
      Y1              =   2610
      Y2              =   2610
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Height          =   195
      Index           =   0
      Left            =   3375
      TabIndex        =   0
      Top             =   225
      Width           =   45
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetTickCount Lib "Kernel32" () As Long
Private mstrTemp As String
Private cMyList  As clsSimpleArray

Private Sub Command1_Click()

  Dim cSA    As clsSimpleArray
  
  Dim lngI  As Long
  Dim lngJ  As Long
  Dim lngSum As Long
  Dim lngMS As Long
  
   Command1.Visible = False
   MousePointer = vbHourglass
   DoEvents
   Set cSA = New clsSimpleArray
   
   '// add
   For lngJ = 1 To 5 '// average 5 samples
      lngMS = GetTickCount
      For lngI = 1 To 32767
         cSA.AddItem CStr(lngI)
      Next lngI
      lngSum = lngSum + (GetTickCount - lngMS)
      If Not (lngJ = 5) Then cSA.Clear
   Next lngJ
   Label1(0).Caption = Format$(lngSum / 5, "0 ms")
   
   
'   '// read
   For lngJ = 1 To 5
      lngMS = GetTickCount
      For lngI = 1 To 32767
         mstrTemp = cSA.List(lngI - 1)
      Next lngI
      lngSum = lngSum + (GetTickCount - lngMS)
   Next lngJ
   Label1(1).Caption = Format$(lngSum / 5, "0 ms")
   
   Set cSA = Nothing
   Command1.Visible = True
   MousePointer = vbDefault

End Sub

Private Sub Command2_Click()
  
  Dim lngI  As Long
  
   Set cMyList = New clsSimpleArray
   cMyList.Sorted = CBool(chkSorted.Value)
   
   cMyList.AddItem "M"
   cMyList.AddItem "G"
   cMyList.AddItem "C"
   cMyList.AddItem "H"
   cMyList.AddItem "G"
   cMyList.AddItem "F"
   cMyList.AddItem "A"
   cMyList.AddItem "C"
   
   List1.Clear
   List2.Clear
   
   For lngI = 0 To cMyList.ListCount - 1
      List1.AddItem cMyList.List(lngI)
   Next lngI
   
   cMyList.RemoveDuplicates
   
   For lngI = 0 To cMyList.ListCount - 1
      List2.AddItem cMyList.List(lngI)
   Next lngI


End Sub

Private Sub Command3_Click()

  Dim udtCol As Collection
  Dim lngI  As Long
  Dim lngJ  As Long
  Dim lngSum As Long
  Dim lngMS As Long
  
   Command3.Visible = False
   MousePointer = vbHourglass
   DoEvents
   Set udtCol = New Collection
   
   '// add
   For lngJ = 1 To 5 '// average 5 samples
      lngMS = GetTickCount
      For lngI = 1 To 32767
         udtCol.Add CStr(lngI)
      Next lngI
      lngSum = lngSum + (GetTickCount - lngMS)
      If Not (lngJ = 5) Then Set udtCol = New Collection
   Next lngJ
   Label1(2).Caption = Format$(lngSum / 5, "0 ms")
   
   
   '// read 1 sample because it takes to long
'   For lngJ = 1 To 2
      lngMS = GetTickCount
      For lngI = 1 To 32767
         mstrTemp = udtCol.Item(lngI)
      Next lngI
      lngSum = lngSum + (GetTickCount - lngMS)
      Label1(3).Caption = CStr(lngSum) & " ms"
'   Next lngJ
'   Label1(3).Caption = format$(lngSum / 2,"0 ms")
   
   Set udtCol = Nothing
   Command3.Visible = True
   MousePointer = vbDefault

End Sub

Private Sub Command4_Click()

  Dim lngI  As Long
  Dim lngJ  As Long
  Dim lngSum As Long
  Dim lngMS As Long
  
   List1.Clear
   List1.Visible = False '// increases speed because graphics don't update
   Command4.Visible = False
   MousePointer = vbHourglass
   DoEvents
   
   '// add
   For lngJ = 1 To 5 '// average 5 samples
      lngMS = GetTickCount
      For lngI = 1 To 32767
         List1.AddItem CStr(lngI)
      Next lngI
      lngSum = lngSum + (GetTickCount - lngMS)
      If Not (lngJ = 5) Then List1.Clear
   Next lngJ
   Label1(4).Caption = Format$(lngSum / 5, "0 ms")
   
   '// read
   For lngJ = 1 To 5
      lngMS = GetTickCount
      For lngI = 1 To 32767
         mstrTemp = List1.List(lngI)
      Next lngI
      lngSum = lngSum + (GetTickCount - lngMS)
   Next lngJ
   Label1(5).Caption = Format$(lngSum / 5, "0 ms")
   
   List1.Clear
   List1.Visible = True
   Command4.Visible = True
   MousePointer = vbDefault
   
End Sub

Private Sub Command5_Click()

  Dim lngI As Long
  
   Set cMyList = New clsSimpleArray
   cMyList.Sorted = CBool(chkSorted.Value)
   
   cMyList.AddItem "M"
   cMyList.AddItem "G"
   cMyList.AddItem "C"
   cMyList.AddItem "H"
   cMyList.AddItem "G"
   cMyList.AddItem "F"
   cMyList.AddItem "A"
   cMyList.AddItem "C"
   
   List1.Clear
   List2.Clear
   
   For lngI = 0 To cMyList.ListCount - 1
      List1.AddItem cMyList.List(lngI)
   Next lngI

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set cMyList = Nothing
    Set Form1 = Nothing
End Sub
