Attribute VB_Name = "D3D11_Shared"
Option Explicit

'=========================================================================
' Shared
'=========================================================================

Public Sub pvArrayLong(aDest() As Long, ParamArray A() As Variant)
    Dim lIdx            As Long
    
    ReDim aDest(0 To UBound(A)) As Long
    For lIdx = 0 To UBound(A)
        aDest(lIdx) = A(lIdx)
    Next
End Sub

Public Sub pvArrayInteger(aDest() As Integer, ParamArray A() As Variant)
    Dim lIdx            As Long
    
    ReDim aDest(0 To UBound(A)) As Integer
    For lIdx = 0 To UBound(A)
        aDest(lIdx) = A(lIdx)
    Next
End Sub

Public Sub pvArraySingle(aDest() As Single, ParamArray A() As Variant)
    Dim lIdx            As Long
    
    ReDim aDest(0 To UBound(A)) As Single
    For lIdx = 0 To UBound(A)
        aDest(lIdx) = A(lIdx)
    Next
End Sub

Public Sub pvInitInputElementDesc(uEntry As D3D11_INPUT_ELEMENT_DESC, uBuffer As UcsBufferType, SemanticName As String, ByVal SemanticIndex As Long, ByVal format As DXGI_FORMAT, ByVal InputSlot As Long, ByVal AlignedByteOffset As Long, ByVal InputSlotClass As D3D11_INPUT_CLASSIFICATION, ByVal InstanceDataStepRate As Long)
    uBuffer.data = StrConv(SemanticName & vbNullChar, vbFromUnicode)
    With uEntry
        .SemanticName = VarPtr(uBuffer.data(0))
        .SemanticIndex = SemanticIndex
        .format = format
        .InputSlot = InputSlot
        .AlignedByteOffset = AlignedByteOffset
        .InputSlotClass = InputSlotClass
        .InstanceDataStepRate = InstanceDataStepRate
    End With
End Sub

Public Sub pvInitViewport(uEntry As D3D11_VIEWPORT, ByVal TopLeftX As Single, ByVal TopLeftY As Single, ByVal Width As Single, ByVal Height As Single, ByVal MinDepth As Single, ByVal MaxDepth As Single)
    With uEntry
        .TopLeftX = TopLeftX
        .TopLeftY = TopLeftY
        .Width = Width
        .Height = Height
        .MinDepth = MinDepth
        .MaxDepth = MaxDepth
    End With
End Sub

