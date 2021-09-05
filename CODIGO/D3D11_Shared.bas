Attribute VB_Name = "D3D11_Shared"
Option Explicit

'=========================================================================
' Shared
'=========================================================================

Public Sub pvArrayLong(aDest() As Long, ParamArray a() As Variant)
    Dim lIdx            As Long
    
    ReDim aDest(0 To UBound(a)) As Long
    For lIdx = 0 To UBound(a)
        aDest(lIdx) = a(lIdx)
    Next
End Sub

Public Sub pvArrayInteger(aDest() As Integer, ParamArray a() As Variant)
    Dim lIdx            As Long
    
    ReDim aDest(0 To UBound(a)) As Integer
    For lIdx = 0 To UBound(a)
        aDest(lIdx) = a(lIdx)
    Next
End Sub

Public Sub pvArraySingle(aDest() As Single, ParamArray a() As Variant)
    Dim lIdx            As Long
    
    ReDim aDest(0 To UBound(a)) As Single
    For lIdx = 0 To UBound(a)
        aDest(lIdx) = a(lIdx)
    Next
End Sub

Public Sub pvInitInputElementDesc(uEntry As D3D11_INPUT_ELEMENT_DESC, uBuffer As UcsBufferType, SemanticName As String, ByVal SemanticIndex As Long, ByVal Format As DXGI_FORMAT, ByVal InputSlot As Long, ByVal AlignedByteOffset As Long, ByVal InputSlotClass As D3D11_INPUT_CLASSIFICATION, ByVal InstanceDataStepRate As Long)
    uBuffer.Data = StrConv(SemanticName & vbNullChar, vbFromUnicode)
    With uEntry
        .SemanticName = VarPtr(uBuffer.Data(0))
        .SemanticIndex = SemanticIndex
        .Format = Format
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

Public Function pvLoadPng(sFilename As String, lWidth As Long, lHeight As Long, lChannels As Long, baData() As Byte) As Boolean
    Const ImageLockModeRead As Long = 1
    Const PixelFormat32bppPARGB As Long = &HE200B
    Dim aInput(0 To 3)  As Long
    Dim hBitmap         As Long
    Dim uData           As BitmapData
    
    If GetModuleHandle("gdiplus") = 0 Then
        aInput(0) = 1
        Call GdiplusStartup(0, aInput(0))
    End If
    If GdipLoadImageFromFile(StrPtr(sFilename), hBitmap) <> 0 Then
        GoTo QH
    End If
    If GdipBitmapLockBits(hBitmap, ByVal 0, ImageLockModeRead, PixelFormat32bppPARGB, uData) <> 0 Then
        GoTo QH
    End If
    lWidth = uData.Width
    lHeight = uData.Height
    lChannels = 4
    ReDim baData(0 To uData.stride * uData.Height - 1) As Byte
    Call CopyMemory(baData(0), ByVal uData.scan0, UBound(baData) + 1)
    '--- success
    pvLoadPng = True
QH:
    If uData.scan0 <> 0 Then
        Call GdipBitmapUnlockBits(hBitmap, uData)
    End If
    If hBitmap <> 0 Then
        Call GdipDisposeImage(hBitmap)
    End If
End Function

