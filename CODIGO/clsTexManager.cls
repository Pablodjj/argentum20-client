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

Public TextureSampler As ID3D11SamplerState

Private Type SURFACE_ENTRY_DYN
    FileName As Integer
    UltimoAcceso As Long
    texture As VBD3D11.ID3D11Texture2D
    Size As Long
    texture_width As Integer
    texture_height As Integer
End Type

Private Type HashNode
    surfaceCount As Integer
    SurfaceEntry() As SURFACE_ENTRY_DYN
End Type

Const HASH_TABLE_SIZE As Long = 337
Private TexList(HASH_TABLE_SIZE - 1) As HashNode

Private mGraphicsDevice              As VBD3D11.ID3D11Device1

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

Public Function GetTexture(ByVal FileName As Integer, ByRef textwidth As Long, ByRef textheight As Long) As VBD3D11.ID3D11Texture2D
    
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

Public Sub Init(ByRef p_GraphicsDevice As VBD3D11.ID3D11Device1, ByVal MaxMemory As Long)
    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero
    'Last Modify Date: 5/04/2005
    '
    '**************************************************************

    Set mGraphicsDevice = p_GraphicsDevice
    mCantidadGraficos = 0
    mFreeMemoryBytes = MaxMemory

End Sub

Private Function CrearGrafico(ByVal Archivo As Integer, ByRef texWidth As Long, ByRef textheight As Long) As Direct3DTexture8

    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero
    'menduz was here
    '
    '**************************************************************
    On Error GoTo ErrHandler
    
    Dim Index         As Integer
    Dim textureHandle As StdPicture, bytArr() As Byte
    
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
    Set TextureSampler = mGraphicsDevice.CreateSamplerState(samplerDesc)
    
    Index = Archivo Mod HASH_TABLE_SIZE
    
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
                
                textureHandle = StdPictureEx.LoadPicture(bytArr)
                
            #Else
                
                '--- Get image location
                Dim PathToFile As String
                PathToFile = App.Path & "\..\Recursos\Graficos\" & LTrim(str(Archivo)) & ".png"
        
                If Not FileExist(PathToFile, vbArchive) Then
                    Debug.Print "¡No se puede cargar el grafico numero " & Archivo & "!"
                    Exit Function

                End If

                '--- Load Image
                '----- Source: https://forums.codeguru.com/showthread.php?256617-How-to-convert-an-IPictureDisp-to-byte-array-(and-the-opposite)
                textureHandle = StdPictureEx.LoadPicture(PathToFile)
                
                ' Extract raw data as a Byte Array from the StdPicture object
                Dim metadata As PropertyBag
                Set metadata = New PropertyBag
                Call metadata.WriteProperty("Image", textureHandle)
                bytArr = metadata.Contents
  
            #End If
            
            '--- Create Texture
            Dim textureDesc             As D3D11_TEXTURE2D_DESC
            Dim textureSubresourceData  As D3D11_SUBRESOURCE_DATA
            
            With textureDesc
                .Width = textureHandle.Width
                .Height = textureHandle.Height
                .MipLevels = 1
                .ArraySize = 1
                .format = DXGI_FORMAT_B8G8R8A8_UNORM_SRGB
                .SampleDesc.count = 1
                .Usage = D3D11_USAGE_IMMUTABLE
                .BindFlags = D3D11_BIND_SHADER_RESOURCE
            End With
            
            With textureSubresourceData
                .pSysMem = VarPtr(bytArr(0))
                .SysMemPitch = LenB(bytArr)
            End With
            
            Set .texture = mGraphicsDevice.CreateTexture2D(textureDesc, textureSubresourceData)
            
            ' Stored in Texlist
            .texture_width = textureHandle.Width
            .texture_height = textureHandle.Height
            .Size = textureSubresourceData.SysMemPitch
            
            ' To be returned by the function
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
