Attribute VB_Name = "modMath"
Public Enum BYTEVALUES
    KiloByte = 1024
    MegaByte = 1048576
    GigaByte = 1073741824
End Enum
Public Function CutDecimal(Number As String, ByPlace As Byte) As String
    Dim Dec As Byte
    Dec = InStr(1, Number, ".", vbBinaryCompare)
    If Dec = 0 Then
        CutDecimal = Number
        Exit Function
    End If
    CutDecimal = Mid(Number, 1, Dec + ByPlace)
End Function


Function GiveByteValues(Bytes As Double) As Double
On Error Resume Next
    If Bytes < BYTEVALUES.KiloByte Then
        GiveByteValues = Bytes & " Bytes"
    ElseIf Bytes >= BYTEVALUES.GigaByte Then
        GiveByteValues = CutDecimal(Bytes / BYTEVALUES.GigaByte, 2)
        what = "Gigabytes"
    ElseIf Bytes >= BYTEVALUES.MegaByte Then
        GiveByteValues = CutDecimal(Bytes / BYTEVALUES.MegaByte, 2)
        what = "Megabytes"
    ElseIf Bytes >= BYTEVALUES.KiloByte Then
        GiveByteValues = CutDecimal(Bytes / BYTEVALUES.KiloByte, 2)
        what = "Kilobytes"
    End If
End Function

