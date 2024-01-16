Attribute VB_Name = "MaatoohLongLong"
Private Declare Function VariantChangeTypeEx Lib "oleaut32" (ByRef pvargDest As Variant, ByRef pvarSrc As Variant, ByVal lcid As Long, ByVal wFlags As Integer, ByVal vt As Integer) As Long
Private Declare Function GetMem4 Lib "msvbvm60" (Src As Any, Dst As Any) As Long
Private Const vbLongLong As Integer = &H14
'

Public Function cLngLng(Expression As Variant) As Variant
    Const LOCALE_INVARIANT As Long = &H7F&
    Dim hr&
    hr& = VariantChangeTypeEx(cLngLng, Expression, LOCALE_INVARIANT, 0, vbLongLong)
    If hr < 0 Then Err.Raise hr
End Function

Public Function HexEx(Number As Variant) As String
    Select Case VarType(Number)
    Case vbLongLong
        Dim Hi&, Lo&
        GetMem4 ByVal VarPtr(Number) + 8, Lo
        GetMem4 ByVal VarPtr(Number) + 12, Hi
        HexEx = Hex$(Hi) & Right$("0000000" & Hex$(Lo), 8)
    Case Else
        HexEx = Hex$(Number)
    End Select
End Function

Public Sub Test()
    Dim v1 As Variant
    Dim v2 As Variant
    v1 = cLngLng("311111111111111116")
    v2 = cLngLng("233333333333333331")
    MsgBox v1 + v2
    Debug.Print HexEx(v1)
    Debug.Print VarType(v1)
End Sub

