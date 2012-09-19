Attribute VB_Name = "Module1"
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public lStartTime As Long

'Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
    

Public Function hex2float(ByVal a As Byte, ByVal b As Byte, ByVal c As Byte, ByVal d As Byte) As Single
Dim k(3)     As Byte     '不应定义为Dim   A(4)   As   Byte，原因为vb的数组下标默认从0开始
Dim Result     As Single

Dim l(3)     As Byte
Dim i     As Integer
k(0) = b
k(1) = a
k(2) = d
k(3) = c

'For i = 0 To 3
'        l(i) = k(3 - i)
'Next
CopyMemory Result, k(0), 4
hex2float = Result
'MsgBox Result
End Function


