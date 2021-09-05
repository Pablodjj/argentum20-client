VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGraphicsEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Engine Context
Private m_d3d11Device           As VBD3D11.ID3D11Device1
Private m_d3d11DeviceContext    As VBD3D11.ID3D11DeviceContext1
Private m_d3d11SwapChain        As VBD3D11.IDXGISwapChain1
Private m_d3d11FrameBufferView  As VBD3D11.ID3D11RenderTargetView

Private m_SpriteBatch           As clsBatch
Private m_TextureManager        As clsTexManager

Private m_vsBlob                As VBD3D11.ID3DBlob
Private m_vertexShader          As VBD3D11.ID3D11VertexShader
Private m_pixelShader           As VBD3D11.ID3D11PixelShader
Private m_inputLayout           As VBD3D11.ID3D11InputLayout
Private m_vertexBuffer          As VBD3D11.ID3D11Buffer
Private m_numVerts              As Long
Private m_stride                As Long
Private m_offset                As Long

Private Type UcsBufferType
    data()              As Byte
End Type

Private Type UcsConstantsType
    modelViewProj       As XMMATRIX
End Type

' Engine States
Private m_isRunning             As Boolean

Public Sub BeginScene(Optional ByRef DestRect As D3D11_RECT, Optional ByVal ColorRGBA As Single = 0)
    
    On Error GoTo Engine_BeginScene_Err
        
    Call m_d3d11DeviceContext.ClearRenderTargetView(m_d3d11FrameBufferView, ColorRGBA)
    
    Dim viewport As D3D11_VIEWPORT
    With viewport
        .Width = DestRect.Right - DestRect.Left
        .Height = DestRect.Bottom - DestRect.Top
        .MaxDepth = 1
    End With
    
    Call m_d3d11DeviceContext.RSSetViewports(1, viewport)
        
    Call m_d3d11DeviceContext.OMSetRenderTargets(1, m_d3d11FrameBufferView, Nothing)

    Call m_d3d11DeviceContext.IASetPrimitiveTopology(D3D11_PRIMITIVE_TOPOLOGY_TRIANGLELIST)
    Call m_d3d11DeviceContext.IASetInputLayout(m_inputLayout)
        
    Call m_d3d11DeviceContext.VSSetShader(m_vertexShader, Nothing, 0)
    m_d3d11DeviceContext.PSSetShader m_pixelShader, Nothing, 0
        
    m_d3d11DeviceContext.IASetVertexBuffers 0, 1, m_vertexBuffer, m_stride, m_offset
    
    Exit Sub

Engine_BeginScene_Err:
    
End Sub

Public Sub EndScene(ByRef DestRect As D3D11_RECT)

On Error GoTo ErrorHandler:
    
    Call SpriteBatch.Flush
    
    
    
ErrorHandler:

        
End Sub

Private Sub Class_Initialize()
        
     '-------------- Create D3D11 Device and Context
    Dim featureLevels() As Long
    Dim creationFlags   As Long
    pvArrayLong featureLevels, D3D_FEATURE_LEVEL_11_0
    creationFlags = D3D11_CREATE_DEVICE_BGRA_SUPPORT
    #If DebugBuild Then
        creationFlags = creationFlags Or D3D11_CREATE_DEVICE_DEBUG
    #End If
    
RetryCreateDevice:
    hResult = D3D11CreateDevice(Nothing, D3D_DRIVER_TYPE_HARDWARE, 0, creationFlags, _
                                featureLevels(0), UBound(featureLevels) + 1, D3D11_SDK_VERSION, _
                                m_d3d11Device, 0, m_d3d11DeviceContext)
    If hResult = DXGI_ERROR_SDK_COMPONENT_MISSING And (creationFlags And D3D11_CREATE_DEVICE_DEBUG) <> 0 Then
        creationFlags = creationFlags And Not D3D11_CREATE_DEVICE_DEBUG
        GoTo RetryCreateDevice
    End If
    If hResult < 0 Then
        Err.Raise hResult, "D3D11CreateDevice"
    End If
    
#If DEBUGGING Then
    '-------------- Set up debug layer to break on D3D11 errors
    Dim d3dDebug        As VBD3D11.ID3D11Debug
    Dim d3dInfoQueue    As VBD3D11.ID3D11InfoQueue
    
    Set d3dDebug = m_d3d11Device
    
    If Not d3dDebug Is Nothing Then
        Set d3dInfoQueue = d3dDebug
        d3dInfoQueue.SetBreakOnSeverity D3D11_MESSAGE_SEVERITY_CORRUPTION, 1
        d3dInfoQueue.SetBreakOnSeverity D3D11_MESSAGE_SEVERITY_ERROR, 1
        Set d3dInfoQueue = Nothing
    End If
    
    Set d3dDebug = Nothing
#End If

    '-------------- Create Swap Chain
    Dim dxgiFactory     As VBD3D11.IDXGIFactory2
    Dim dxgiDevice      As VBD3D11.IDXGIDevice1
    Dim dxgiAdapter     As VBD3D11.IDXGIAdapter
    Dim adapterDesc     As VBD3D11.DXGI_ADAPTER_DESC
    Dim d3d11SwapChainDesc As VBD3D11.DXGI_SWAP_CHAIN_DESC1
    
    Set dxgiDevice = m_d3d11Device
    Set dxgiAdapter = dxgiDevice.GetAdapter()
    Set dxgiDevice = Nothing
    
    dxgiAdapter.GetDesc adapterDesc
    Debug.Print "Graphics Device: " & Replace(adapterDesc.Description, vbNullChar, vbNullString)
    
    Set dxgiFactory = dxgiAdapter.GetParent(IIDFromString(szIID_IDXGIFactory2))
    With d3d11SwapChainDesc
        .Width = 0 '--- use window width
        .Height = 0 '--- use window height
        .format = DXGI_FORMAT_B8G8R8A8_UNORM_SRGB
        .SampleDesc.count = 1
        .SampleDesc.Quality = 0
        .BufferUsage = DXGI_USAGE_RENDER_TARGET_OUTPUT
        .BufferCount = 2
        .Scaling = DXGI_SCALING_STRETCH
        .SwapEffect = DXGI_SWAP_EFFECT_DISCARD
        .AlphaMode = DXGI_ALPHA_MODE_UNSPECIFIED
        .flags = 0
    End With
    Set m_d3d11SwapChain = dxgiFactory.CreateSwapChainForHwnd(m_d3d11Device, frmMain.hWnd, d3d11SwapChainDesc, ByVal 0, Nothing)
    
    '-------------- Create Framebuffer Render Target
    Dim d3d11FrameBuffer As ID3D11Texture2D
    Set d3d11FrameBuffer = m_d3d11SwapChain.GetBuffer(0, IIDFromString(szIID_ID3D11Texture2D))
    Set m_d3d11FrameBufferView = m_d3d11Device.CreateRenderTargetView(d3d11FrameBuffer, ByVal 0)
    
    '-------------- Create Draw Pipeline
    Call Init_Pipeline
    
    '-------------- Initialize Vertex Buffer
    Set m_SpriteBatch = New clsBatch
    Call m_SpriteBatch.Initialize(2000)
    
    '-------------- Initialize Surface Handler
    Set m_TextureManager = New clsTexManager
    Call m_TextureManager.Init(m_d3d11Device, 2000)
    
End Sub

Private Sub Init_Pipeline()
    
    '--- Create Vertex Shader
    Dim shaderCompileErrorsBlob As ID3DBlob
    Dim errorString     As String
    hResult = D3DCompileFromFile(App.Path & "\CODIGO\shaders.hlsl", ByVal 0, ByVal 0, "vs_main", "vs_5_0", 0, 0, m_vsBlob, shaderCompileErrorsBlob)
    If hResult < 0 Then
        If hResult = LNG_FACILITY_WIN32 Or ERROR_FILE_NOT_FOUND Then
            errorString = "Could not compile shader; file not found"
        ElseIf Not shaderCompileErrorsBlob Is Nothing Then
            errorString = pvToString(shaderCompileErrorsBlob.GetBufferPointer())
        End If
        MsgBox errorString, vbCritical, "Shader Compiler Error"
        Unload Me
        GoTo QH
    End If
    Set m_vertexShader = m_d3d11Device.CreateVertexShader(m_vsBlob.GetBufferPointer(), m_vsBlob.GetBufferSize(), Nothing)
    
    '--- Create Pixel Shader
    Dim psBlob As ID3DBlob
    hResult = D3DCompileFromFile(App.Path & "\CODIGO\shaders.hlsl", ByVal 0, ByVal 0, "ps_main", "ps_5_0", 0, 0, psBlob, shaderCompileErrorsBlob)
    If hResult < 0 Then
        If hResult = LNG_FACILITY_WIN32 Or ERROR_FILE_NOT_FOUND Then
            errorString = "Could not compile shader; file not found"
        ElseIf Not shaderCompileErrorsBlob Is Nothing Then
            errorString = pvToString(shaderCompileErrorsBlob.GetBufferPointer())
        End If
        MsgBox errorString, vbCritical, "Shader Compiler Error"
        Unload Me
        GoTo QH
    End If
    Set m_pixelShader = m_d3d11Device.CreatePixelShader(psBlob.GetBufferPointer(), psBlob.GetBufferSize(), Nothing)
    Set psBlob = Nothing
    
     '--- Create Input Layout
    Dim inputElementDesc(0 To 2)  As D3D11_INPUT_ELEMENT_DESC
    Dim nameBuffer(0 To 2) As UcsBufferType
    pvInitInputElementDesc inputElementDesc(0), nameBuffer(0), "POSITION", 0, DXGI_FORMAT_R32G32_FLOAT, 0, 0, D3D11_INPUT_PER_VERTEX_DATA, 0
    pvInitInputElementDesc inputElementDesc(1), nameBuffer(1), "COLOR", 0, DXGI_FORMAT_R32G32B32A32_FLOAT, 1, D3D11_APPEND_ALIGNED_ELEMENT, D3D11_INPUT_PER_VERTEX_DATA, 0
    pvInitInputElementDesc inputElementDesc(1), nameBuffer(1), "TEXCOORD", 0, DXGI_FORMAT_R32G32_FLOAT, 2, 0, D3D11_INPUT_PER_VERTEX_DATA, 0
    Set m_inputLayout = m_d3d11Device.CreateInputLayout(inputElementDesc(0), UBound(inputElementDesc) + 1, m_vsBlob.GetBufferPointer(), m_vsBlob.GetBufferSize())
    
End Sub

Private Sub InitConstants()
    
    
End Sub

Private Sub Class_Terminate()
    
    Set m_d3d11Device = Nothing
    Set m_d3d11DeviceContext = Nothing
    Set m_d3d11SwapChain = Nothing
    Set m_d3d11FrameBufferView = Nothing
    
    Set m_Batch = Nothing
    
End Sub
