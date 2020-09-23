VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit


Private m_TestType() As tTestType

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function VarPtrArray Lib "msvbvm50.dll" Alias "VarPtr" (ptr() As Any) As Long


Public Property Get DataTestPtr() As Long
   'Returns pointer to the private m_TestType() array
    CopyMemory DataTestPtr, ByVal VarPtrArray(m_TestType), 4&
End Property


Private Sub Class_Initialize()
    Dim i As Long

   'Some initial values
    ReDim m_TestType(0)
    With m_TestType(0)
        .Integer1 = 1010
        .Integer2 = 2020
        .Long1 = 11111
        .Long2 = 22222
        .String1 = "String One"
        .String2 = "String Two"
        With .SubType
            .datas = 10
            ReDim .data(.datas - 1)
            For i = 0 To .datas - 1
                .data(i) = i
            Next i
        End With
        ReDim .LongArray(7)
        For i = 0 To 7
            .LongArray(i) = i * 1000
        Next i
        .StartDate = Now
    End With

End Sub
