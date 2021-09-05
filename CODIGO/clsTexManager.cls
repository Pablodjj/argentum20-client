VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTexManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const HASH_TABLE_SIZE As Long = 337

Public TextureSampler As ID3D11SamplerState

Private Type SURFACE_ENTRY_DYN

    FileName As Integer
    UltimoAcceso As Long
    texture As ID3D11Texture2D
    Size As Long
    texture_width As Integer
    texture_height As Integer

End Type

Private Type HashNode

    surfaceCount As Integer
    SurfaceEntry() As SURFACE_ENTRY_DYN

End Type

Private TexList(HASH_TABLE_SIZE - 1) As HashNode

Private mD3D                         As D3DX8

Private device                       As Direct3DDevice8

Private mMaxEntries                  As Integer

Private mCantidadGraficos            As Integer

Private mFreeMemoryBytes             As Long

Private Sub Class_Terminate()
    
    On Error GoTo Class_Terminate_Err
    

    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero
    'Last Modify Date: 5/04/2005
    '
    '**************************************************************
    Dim I As Long

    Dim J As Long
    
    'Destroy every surface in memory
    For I = 0 To HASH_TABLE_SIZE - 1

        With TexList(I)

            For J = 1 To .surfaceCount
                Set .SurfaceEntry(J).texture = Nothing
            Next J
            
            'Destroy the arrays
            Erase .SurfaceEntry

        End With

    Next I

    
    Exit Sub

Class_Terminate_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsTexManager.Class_Terminate", Erl)
    Resume Next
    
End Sub

Public Function GetTexture(ByVal FileName As Integer, ByRef textwidth As Long, ByRef textheight As Long) As Direct3DTexture8
    
    On Error GoTo GetTexture_Err
    

    If FileName = 0 Then
        Debug.Print "0 GRH ATMPT TO BE LOADED"
        Exit Function

    End If

    Dim I As Long

    ' Search the index on the list
    With TexList(FileName Mod HASH_TABLE_SIZE)

        For I = 1 To .surfaceCount

            If .SurfaceEntry(I).FileName = FileName Then
                .SurfaceEntry(I).UltimoAcceso = GetTickCount()
                textwidth = .SurfaceEntry(I).texture_width
                textheight = .SurfaceEntry(I).texture_height
                Set GetTexture = .SurfaceEntry(I).texture
                Exit Function

            End If

        Next I

    End With

    'Not in memory, load it!
    Set GetTexture = CrearGrafico(FileName, textwidth, textheight)
    
    Exit Function

GetTexture_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsTexManager.GetTexture", Erl)
    Resume Next
    
End Function

Public Function Init(ByRef D3D8 As D3DX8, ByRef d3d_device As Direct3DDevice8, ByVal MaxMemory As Long) As Boolean
    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero
    'Last Modify Date: 5/04/2005
    '
    '**************************************************************
    
    On Error GoTo Init_Err
    
    mCantidadGraficos = 0
    
    'Seteamos el objeto
    Set mD3D = D3D8
    Set device = d3d_device
    mFreeMemoryBytes = MaxMemory
    
    Init = True

    
    Exit Function

Init_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsTexManager.Init", Erl)
    Resume Next
    
End Function

Private Function CrearGrafico(ByVal Archivo As Integer, ByRef texWidth As Long, ByRef textheight As Long) As Direct3DTexture8

    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero
    'menduz was here
    '
    '**************************************************************
    On Error GoTo ErrHandler

    Dim surface_desc As D3DSURFACE_DESC
    Dim Index        As Integer
    Dim DirArchico   As String
    Dim bytArr()     As Byte
    
    Index = Archivo Mod HASH_TABLE_SIZE
    
    '--- Create Sampler State
    Dim samplerDesc As D3D11_SAMPLER_DESC
    With samplerDesc
        .Filter = D3D11_FILTER_MIN_MAG_MIP_POINT
        .AddressU = D3D11_TEXTURE_ADDRESS_BORDER
        .AddressV = D3D11_TEXTURE_ADDRESS_BORDER
        .AddressW = D3D11_TEXTURE_ADDRESS_BORDER
        .BorderColor(0) = 1!
        .BorderColor(1) = 1!
        .BorderColor(2) = 1!
        .BorderColor(3) = 1!
        .ComparisonFunc = D3D11_COMPARISON_NEVER
    End With
    Set m_samplerState = m_d3d11Device.CreateSamplerState(samplerDesc)
    
    '--- Load Image
    Dim texWidth         As Long
    Dim texHeight        As Long
    Dim texNumChannels   As Long
    Dim testTextureBytes() As Byte
    Dim texBytesPerRow   As Long
    If Not pvLoadPng(PathCombine(App.Path, "testTexture.png"), texWidth, texHeight, texNumChannels, testTextureBytes) Then
        MsgBox "Error loading testTexture.png", vbExclamation, "Load Image"
        Unload Me
        GoTo QH
    End If
    texBytesPerRow = texWidth * texNumChannels
    
    '--- Create Texture
    Dim textureDesc     As D3D11_TEXTURE2D_DESC
    Dim textureSubresourceData As D3D11_SUBRESOURCE_DATA
    Dim texture         As ID3D11Texture2D
    With textureDesc
        .Width = texWidth
        .Height = texHeight
        .MipLevels = 1
        .ArraySize = 1
        .Format = DXGI_FORMAT_B8G8R8A8_UNORM_SRGB
        .SampleDesc.count = 1
        .Usage = D3D11_USAGE_IMMUTABLE
        .BindFlags = D3D11_BIND_SHADER_RESOURCE
    End With
    With textureSubresourceData
        .pSysMem = VarPtr(testTextureBytes(0))
        .SysMemPitch = 4
    End With
    Set texture = m_d3d11Device.CreateTexture2D(textureDesc, textureSubresourceData)

    
    With TexList(Index)
        .surfaceCount = .surfaceCount + 1
        ReDim Preserve .SurfaceEntry(1 To .surfaceCount) As SURFACE_ENTRY_DYN

        With .SurfaceEntry(.surfaceCount)
            'Nombre
            .FileName = Archivo
            
            'Ultimo acceso
            .UltimoAcceso = GetTickCount()

            #If Compresion = 1 Then

                If Not Extract_File_To_Memory(Graphics, App.Path & "\..\Recursos\OUTPUT\", LTrim(str(Archivo)) & ".png", bytArr, ResourcesPassword) Then
                    Debug.Print "¡No se puede cargar el grafico numero " & Archivo & "!"
                    Exit Function

                End If
                
                Dim textureFile As StdPicture
                textureFile = StdPictureEx.LoadPicture(bytArr)
                
                'Set .texture = mD3D.CreateTextureFromFileInMemoryEx(device, bytArr(0), UBound(bytArr) + 1, D3DX_DEFAULT, D3DX_DEFAULT, 1, 0, D3DFMT_A8R8G8B8, D3DPOOL_MANAGED, D3DX_FILTER_NONE, D3DX_FILTER_NONE, &HFF000000, ByVal 0, ByVal 0)
            #Else

                Dim PathToFile As String

                PathToFile = App.Path & "\..\Recursos\Graficos\" & LTrim(str(Archivo)) & ".png"
        
                If Not FileExist(PathToFile, vbArchive) Then
                    Debug.Print "¡No se puede cargar el grafico numero " & Archivo & "!"
                    Exit Function

                End If

                'Set .texture = mD3D.CreateTextureFromFileEx(device, PathToFile, D3DX_DEFAULT, D3DX_DEFAULT, 1, 0, D3DFMT_A8R8G8B8, D3DPOOL_MANAGED, D3DX_FILTER_NONE, D3DX_FILTER_NONE, &HFF000000, ByVal 0, ByVal 0)
            #End If
                
            .texture.GetLevelDesc 0, surface_desc
            .texture_width = surface_desc.Width
            .texture_height = surface_desc.Height
            .Size = surface_desc.Size
            texWidth = .texture_width
            textheight = .texture_height
            
            Set CrearGrafico = TexList(Index).SurfaceEntry(TexList(Index).surfaceCount).texture

        End With

    End With

    Exit Function

ErrHandler:
    Debug.Print "ERROR EN GRHLOAD>" & Archivo & ".png"

End Function

Public Sub Grh_Load(ByVal file_num As Integer)
    
    On Error GoTo Grh_Load_Err
    
    Static d3dTextures As D3D8Textures

    '**************************************************************
    'Author: Augusto José Rando
    'Last Modify Date: 7/16/2005
    'We don't return nothing
    '**************************************************************
    Set d3dTextures.texture = CrearGrafico(file_num, 0, 0)
    
    
    Exit Sub

Grh_Load_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsTexManager.Grh_Load", Erl)
    Resume Next
    
End Sub

