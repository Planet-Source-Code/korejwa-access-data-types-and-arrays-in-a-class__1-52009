Attribute VB_Name = "Module1"

Option Explicit


'A complex variable Type to be used within a Class

Public Type tTestSubType
    data() As Byte
    datas As Long
End Type

Public Type tTestType
    Integer1     As Integer
    Long1        As Long
    String1      As String
    SubType      As tTestSubType
    LongArray()  As Long
    Integer2     As Integer
    Long2        As Long
    String2      As String
    StartDate    As Date
End Type
