VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3780
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   ScaleHeight     =   3780
   ScaleWidth      =   6135
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Access arrays and data types within a Class"
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   5895
      Begin VB.CommandButton Command1 
         Caption         =   "Access Variable Type in Class"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   2880
         Width           =   5655
      End
      Begin VB.ListBox List1 
         Height          =   2400
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   5655
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Private m_Test As New cTest

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function VarPtrArray Lib "msvbvm50.dll" Alias "VarPtr" (ptr() As Any) As Long


Private Sub Command1_Click()
    Dim i As Long
    Dim Temp() As tTestType


   'Copy array pointer from the class to Temp()
    CopyMemory ByVal VarPtrArray(Temp), m_Test.DataTestPtr, 4&

   'Show values
    With Temp(0)
        List1.AddItem ".Integer1" & vbTab & "= " & CStr(.Integer1)
        List1.AddItem ".Integer2" & vbTab & "= " & CStr(.Integer2)
        List1.AddItem ".Long1" & vbTab & "= " & CStr(.Long1)
        List1.AddItem ".Long2" & vbTab & "= " & CStr(.Long2)
        List1.AddItem ".String1" & vbTab & "= " & .String1
        List1.AddItem ".String2" & vbTab & "= " & .String2
        List1.AddItem ".SubType.datas" & vbTab & "= " & CStr(.SubType.datas)
        For i = LBound(.SubType.data) To UBound(.SubType.data)
            List1.AddItem "  " & CStr(i) & vbTab & "= " & CStr(.SubType.data(i))
        Next i
        For i = LBound(.LongArray) To UBound(.LongArray)
            List1.AddItem "  Long Array " & CStr(i) & vbTab & "= " & CStr(.LongArray(i))
        Next i
        List1.AddItem "Class Initialized " & vbTab & "= " & CStr(.StartDate)
    End With
   'Note:  You can also assign and change values in the class

   'Clear the Temp() array pointer
    CopyMemory ByVal VarPtrArray(Temp), 0&, 4&
   'Warning!  VB IDE can crash if the code stops while Temp() still points to the data type

End Sub
