Attribute VB_Name = "Engine_Input"
Option Explicit

' Fuente: http://directx4vb.vbgamer.com/DirectX4VB/Tutorials/DirectX8/IN_Lesson01.asp

Public DirectInput As DirectInput8
Public DI_Device As DirectInputDevice8

Private DI_State As DIKEYBOARDSTATE
Public KeyState(0 To 255) As Boolean       'so we can detect if the key has gone up or down!

Public Const DI_BufferSize As Long = 10    'how many events the buffer holds.
                                           'This can be 1 if using event based, but 10-20 if polling based...

Public Sub Init_InputDevice()
    
    On Error GoTo ErrorHandler:

    '//0. Any variables
     Dim I As Long
     Dim DevProp As DIPROPLONG
     Dim DevInfo As DirectInputDeviceInstance8
     Dim pBuffer(0 To BufferSize) As DIDEVICEOBJECTDATA

    ' Initialize required objects
    If DirectX Is Nothing Then Set DirectX = New DirectX8
    
    Set DirectInput = DirectX.DirectInputCreate
    Set DirectInputDevice = DI.CreateDevice("GUID_SysKeyboard") 'the string is important, not just a random string..
    
    ' Setup
    Call DirectInputDevice.SetCommonDataFormat(DIFORMAT_KEYBOARD)
    Call DirectInputDevice.SetCooperativeLevel(frmMain.hwnd, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE)
                                                             
    ' Set up the buffer...
    DevProp.lHow = DIPH_DEVICE
    DevProp.lData = BufferSize
    Call DirectInputDevice.SetProperty(DIPROP_BUFFERSIZE, DevProp)
                                                             
    'let DirectX know that we want to use the device now.
    Call DirectInputDevice.Acquire
    
    Exit Sub

ErrorHandler:
    
    Call MsgBox("Ocurrió un error al inicializar el motor gráfico. Cod: I500", vbCritical)
    Call CloseClient
    
End Sub

Public Function Engine_ProcessInput() As Byte
    
    'a. retrieve the information
    Call DirectInputDevice.GetDeviceStateKeyboard(DI_State)  'get the keyboard state

    On Error Resume Next 'ignore the prev. err handler

    Call DIDevice.GetDeviceData(pBuffer, DIGDD_DEFAULT) 'retrieve buffer info.

    If Err.Number = DI_BUFFEROVERFLOW Then 'check for an error..
        Debug.Print vbCr & "BUFFER OVERFLOW (Compensating)..."
        Engine_ProcessInput = -1 'too much data, just loop around to the next loop..

    End If

    On Error GoTo ErrorHandler: 'reinstate the old error handler.

    'b. sort through this data...
    'most apps would look at a specific key rather
    'than loop through them all; but we're interested in all of them
    For I = 0 To 255 'loop through all the keys

        If DI_State.key(I) = 128 And (Not KeyState(I) = True) Then 'it's been pressed...
            'the value will almost always be 128, indicating a key was pressed...
            KeyState(I) = True
        End If

    Next I
                                                             
    'c. check for any key_up events
    For I = 0 To BufferSize

        If KeyState(pBuffer(I).lOfs) = True And pBuffer(I).lData = 0 Then
            KeyState(pBuffer(I).lOfs) = False
        End If

    Next I
    
    Exit Function
    
ErrorHandler:
End Function
