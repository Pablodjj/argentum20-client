VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSoundEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*****************************************************************
'clsSoundEngine - ImperiumAO - v1.3.0
'
'
'*****************************************************************
'Respective portions copyrighted by contributors listed below.
'
'This library is free software; you can redistribute it and/or
'modify it under the terms of the GNU Lesser General Public
'License as published by the Free Software Foundation version 2.1 of
'the License
'
'This library is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
'Lesser General Public License for more details.
'
'You should have received a copy of the GNU Lesser General Public
'License along with this library; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'*****************************************************************

'*****************************************************************
'Augusto Jos� Rando (barrin@imperiumao.com.ar)
'   - First Relase
'*****************************************************************

Option Explicit

Private dX                        As DirectX8

'**** Direct Sound ********
Private DS                        As DirectSound8

Private DS_Enum                   As DirectSoundEnum8

'**** Direct Music ********
Private dmLoader                  As DirectMusicLoader8

Private dmPerf                    As DirectMusicPerformance8

Private dmSeg                     As DirectMusicSegment8

Private dmSegState                As DirectMusicSegmentState8

Private dmPath                    As DirectMusicAudioPath8

'**** MP3 ********
Dim IMC                           As IMediaControl

Dim IBA                           As IBasicAudio

Dim IME                           As IMediaEvent

Dim IMPos                         As IMediaPosition

'**** AMBIENT ********
Dim ambient_buffer                As DirectSoundSecondaryBuffer8

Private mAmbienteActual           As Integer

Private mLastAmbienteActual       As Integer

Private mMusActual                As Integer

Private WavPath                   As String

Private MidiPath                  As String

Private MP3Path                   As String

'Manejador de buffers
Private Buffers                   As CBufferMan

'El vol�men actual real (sonidos)
Private mEffectsVolume            As Long

'El vol�men actual real (ambientes)
Private mEffectsVolumeAmbient     As Long

'El vol�men actual real (m�sica)
Private mMusicVolume              As Long

'El vol�men m�ximo en la configuraci�n (sonidos)
Private mEffectsVolumeMax         As Long

'El vol�men m�ximo en la configuraci�n (ambientes)
Private mAmbientVolumeMax         As Long

'El vol�men m�ximo en la configuraci�n (m�sica)
Private mMusicVolumeMax           As Long

'�Estamos haciendo efecto fade?
Private mFadingStatus             As Byte

Private mFadingMusicMod           As Long

'�Pr�ximo MP3 o MIDI?
Private mNextMusic                As Integer

Private INTERIOR_VOLUME           As Long

Private EXTERIOR_VOLUME           As Long

Private TimeElapsed               As Long

Private mInvert_Sound             As Boolean

Private Const VOLUME_DISTANCE_MOD As Long = 200

Private lastMid                   As Boolean

Private Property Let SetWavPath(ByVal inString As String)
    WavPath = inString

End Property

Private Property Let SetMP3Path(ByVal inString As String)
    MP3Path = inString

End Property

Private Property Let SetMidiPath(ByVal inString As String)
    MidiPath = inString

End Property

Public Function Engine_DeInitialize()

    If Not Buffers Is Nothing Then
        Call Buffers.BorrarTodo
        Set Buffers = Nothing

    End If

    'stop music playback
    If Not (dmPerf Is Nothing) Then
        ' dmPerf.StopEx dmSeg, 0, 0
        dmPerf.CloseDown

    End If
    
    'delete DirectMusic objects
    Set dmLoader = Nothing
    Set dmSeg = Nothing
    Set dmPath = Nothing
    Set dmPerf = Nothing
    Set dX = Nothing

    Call Music_Empty

    Set dmSegState = Nothing

    If Not dmPath Is Nothing Then
        Call dmPath.Activate(False)
        Set dmPath = Nothing

    End If

    If Not dmPerf Is Nothing Then
        dmPerf.CloseDown
        Set dmPerf = Nothing

    End If

    Set dmLoader = Nothing
    Set dmSeg = Nothing

    Set DS_Enum = Nothing
    Set DS = Nothing
    Set dX = Nothing

End Function

Public Function Initialize_Engine(ByVal main_hwnd As String, ByVal Set_Wav_Path As String, ByVal Set_MP3_Path As String, ByVal Set_Midi_Path As String, Optional ByVal Play_only_on_focus As Boolean = True, Optional ByVal Initialize_Sounds As Boolean = True, Optional ByVal Initialize_Music As Boolean = True, Optional ByVal sound_volume As Long, Optional ByVal music_volume As Long, Optional ByVal invert_snd As Boolean) As Boolean

    On Error GoTo errhandler

    Dim dmA As DMUS_AUDIOPARAMS

    Set dX = New DirectX8

    If Initialize_Sounds Then
        SetWavPath = Set_Wav_Path
        Set DS_Enum = dX.GetDSEnum
        Set DS = dX.DirectSoundCreate(DS_Enum.GetGuid(1))
    
        Set Buffers = New CBufferMan
        Buffers.Path = Set_Wav_Path
        Call Buffers.Init(DS, 400)
    
        If Play_only_on_focus Then
            DS.SetCooperativeLevel main_hwnd, DSSCL_NORMAL
        Else
            DS.SetCooperativeLevel main_hwnd, DSSCL_PRIORITY

        End If

        VolumenActual = sound_volume
        mInvert_Sound = invert_snd

    End If

    If Initialize_Music Then
        SetMidiPath = Set_Midi_Path
        SetMP3Path = Set_MP3_Path
        VolumenActualMusicMax = music_volume
        Set dmLoader = dX.DirectMusicLoaderCreate
        Set dmPerf = dX.DirectMusicPerformanceCreate
        dmPerf.InitAudio main_hwnd, DMUS_AUDIOF_ENVIRON + DMUS_AUDIOF_BUFFERS + DMUS_AUDIOF_3D + DMUS_AUDIOF_EAX + DMUS_AUDIOF_DMOS, dmA
        dmPerf.SetMasterVolume (music_volume)
        Set dmPath = dmPerf.CreateStandardAudioPath(DMUS_APATH_DYNAMIC_3D, 64, True)

    End If

    Initialize_Engine = True

    Exit Function

errhandler:
    Initialize_Engine = False

End Function

Public Property Get AmbienteActual() As Integer
    AmbienteActual = mAmbienteActual

End Property

Public Property Let AmbienteActual(ByVal vNewValue As Integer)

    'If (Not Meteo_Engine.SecondaryStatus = 2) Or (CurrentUser.MapExt = 0) Then
    mAmbienteActual = vNewValue
    Call Ambient_Stop
    Call Ambient_Load(mAmbienteActual, VolAmbient)
    Call Ambient_Play
    'Else
    LastAmbienteActual = vNewValue
    'End If

End Property

Public Property Get LastAmbienteActual() As Integer
    LastAmbienteActual = mLastAmbienteActual

End Property

Public Property Let LastAmbienteActual(ByVal vNewValue As Integer)
    mLastAmbienteActual = vNewValue

End Property

Public Property Get MusicActual() As Integer
    MusicActual = mMusActual

End Property

Public Property Let MusicActual(ByVal vNewValue As Integer)
    mMusActual = vNewValue

End Property

Public Property Let Fading(ByVal vNewValue As Long)
    mFadingMusicMod = vNewValue
    TimeElapsed = (GetTickCount() And &H7FFFFFFF)

End Property

Public Property Get Fading() As Long
    Fading = mFadingMusicMod

End Property

Public Property Get VolumenActual() As Long
    VolumenActual = mEffectsVolumeMax

End Property

Public Property Let VolumenActual(ByVal vNewValue As Long)

    On Error Resume Next

    If vNewValue > 0 Then vNewValue = 0
    If vNewValue < -4000 Then vNewValue = -4000

    mEffectsVolumeMax = vNewValue
    Buffers.ModificarVolumenTodo (vNewValue)
    EXTERIOR_VOLUME = CLng((vNewValue - 1) * 20)
    INTERIOR_VOLUME = CLng((vNewValue - 1) * 45)

    mEffectsVolumeAmbient = -4000

End Property

Public Property Get InvertirSonido() As Boolean
    InvertirSonido = mInvert_Sound

End Property

Public Property Let InvertirSonido(ByVal vNewValue As Boolean)
    mInvert_Sound = vNewValue

End Property

Public Property Let VolumenActualMusicMax(ByVal vNewValue As Long)

    If vNewValue > 0 Then vNewValue = 0
    If vNewValue < -4000 Then vNewValue = -4000

    mMusicVolumeMax = vNewValue

End Property

Public Property Get VolumenActualMusicMax() As Long
    VolumenActualMusicMax = mMusicVolumeMax

End Property

Public Property Get VolumenActualMusic() As Long
    VolumenActualMusic = mMusicVolume

End Property

Public Property Let VolumenActualMusic(ByVal vNewValue As Long)
    mMusicVolume = vNewValue

End Property

Public Property Let NextMusic(ByVal vNewValue As Integer)

    If mFadingStatus <> 2 And vNewValue <> 0 Then 'Si no estamos en fade up ya...
        mNextMusic = vNewValue
    Else
        mNextMusic = vNewValue
        mFadingStatus = 0

    End If

End Property

Public Function Sound_Stop_All() As Boolean

    Buffers.DetenerTodo
    mAmbienteActual = 0
    Sound_Stop_All = True

End Function

Public Function Sound_Play(ByVal Sound_File As Integer, Optional ByVal Do_Loop As Boolean = False, Optional ByVal Volume As Long = 0, Optional ByVal pan As Long = 0) As Boolean

    Dim Sound_Buffer As DirectSoundSecondaryBuffer8

    On Error GoTo Error_Handl
    
    If (Sound_File <= 0) Or (fX = 0) Then Exit Function
    
    If Not Buffers Is Nothing Then
        If Volume = 0 Then Volume = mEffectsVolumeMax
        Call Buffers.Reproducir(Sound_File, IIf(Do_Loop, DSBPLAY_LOOPING, DSBPLAY_DEFAULT), Volume, pan)
        Sound_Play = True

    End If
    
    Exit Function
    
Error_Handl:
    Sound_Play = False

End Function

Public Function Sound_Stop(ByVal Sound_File As Integer) As Boolean

    On Error GoTo Error_Handl
    
    If Sound_File <= 0 Then Exit Function
    
    If Not Buffers Is Nothing Then
        Call Buffers.Detener(Sound_File)
        Sound_Stop = True

    End If
        
    Exit Function

Error_Handl:
    Sound_Stop = False

End Function

Public Function Music_Empty() As Boolean
    
    On Error GoTo Error_Handl
    
    Music_Empty = True
    Exit Function
    
    If ObjPtr(IMC) > 0 Then
        IMC.Stop

    End If
    
    Set IBA = Nothing
    Set IME = Nothing
    Set IMPos = Nothing
    Set IMC = Nothing
    
    Music_Empty = True
    Exit Function

Error_Handl:
    Music_Empty = False

End Function

Public Function Music_Load(ByVal file_str As String, Optional ByVal Volume As Long = 0, Optional ByVal balance As Long = 0) As Boolean '**** Loads a MP3 *****
    '**************************************************************
    'Author: Augusto Jos� Rando (barrin@imperiumao.com.ar)
    'Last Modify Date: 2/08/2006
    '**************************************************************
    
    Music = CONST_MIDI

    On Error GoTo Error_Handl
    
    Dim ret As Long
    
    If (LenB(file_str) <= 0) Or (Musica = CONST_DESHABILITADA) Then Exit Function
 
    #If Compresion = 1 Then

        If Music = CONST_MP3 Then
            If Extract_File(mp3, App.Path & "\..\Recursos\OUTPUT\", file_str & ".mp3", Windows_Temp_Dir, False) Then
                If Not Music_Empty() = True Then GoTo Error_Handl
                Set IMC = New FilgraphManager
                IMC.RenderFile Windows_Temp_Dir & file_str & ".mp3"
                Kill Windows_Temp_Dir & file_str & ".mp3"
                lastMid = False
            ElseIf Extract_File(midi, MidiPath, file_str & ".mid", Windows_Temp_Dir, False) Then

                If Not Music_Empty() = True Then GoTo Error_Handl
                Set dmSeg = dmLoader.LoadSegment(Windows_Temp_Dir & file_str & ".mid")
                dmSeg.SetStandardMidiFile
                dmSeg.Download dmPath
                lastMid = True
            Else
                GoTo Error_Handl

            End If

        ElseIf Music = CONST_MIDI Then

            If Extract_File(midi, App.Path & "\..\Recursos\OUTPUT\", file_str & ".mid", Windows_Temp_Dir, False) Then
                If Not Music_Empty() = True Then GoTo Error_Handl
                Set dmSeg = dmLoader.LoadSegment(Windows_Temp_Dir & file_str & ".mid")
                dmSeg.SetStandardMidiFile
                dmSeg.Download dmPath
                lastMid = True

            End If

        Else
            GoTo Error_Handl

        End If
        
    #Else

        Dim PathToFile As String
    
        If Music = CONST_MP3 Then
            PathToFile = App.Path & "/../Recursos/MP3/" & file_str & ".mp3"

            If FileExist(PathToFile, vbArchive) Then
                If Not Music_Empty() = True Then GoTo Error_Handl
                Set IMC = New FilgraphManager
                IMC.RenderFile PathToFile
                lastMid = False

            End If

        ElseIf Music = CONST_MIDI Then
            PathToFile = App.Path & "/../Recursos/midi/" & file_str & ".mid"

            If FileExist(PathToFile, vbArchive) Then
                If Not Music_Empty() = True Then GoTo Error_Handl
                Set dmSeg = dmLoader.LoadSegment(PathToFile)
                dmSeg.SetStandardMidiFile
                dmSeg.Download dmPath
                lastMid = True

            End If
    
        Else
            GoTo Error_Handl

        End If
        
    #End If
    
    If Volume < -4000 Then Volume = -4000
    If balance > 5000 Then balance = 5000
    If balance < -5000 Then balance = -5000
    
    If lastMid = False Then
        Set IBA = IMC
        IBA.Volume = Volume
        IBA.balance = balance
        
        Set IME = IMC
        
        Set IMPos = IMC

        If ObjPtr(IMPos) Then IMPos.Rate = 1#
        IMPos.CurrentPosition = 0
    Else
        Call dmSeg.SetRepeats(-1)

    End If
    
    mMusActual = file_str
    Music_Load = True
    Exit Function

Error_Handl:
    Debug.Print "error al cargar musica"
    Music_Load = False

End Function

Public Function Music_Play() As Boolean
    '**************************************************************
    'Author: Augusto Jos� Rando (barrin@imperiumao.com.ar)
    'Last Modify Date: 2/08/2006
    '**************************************************************
    
    On Error GoTo Error_Handl
    
    Dim ret As Long
    
    If (Musica = CONST_DESHABILITADA) Then Exit Function
    
    If lastMid = False Then
        IMC.Run
    Else
        Set dmSegState = dmPerf.PlaySegmentEx(dmSeg, 0, 0, Nothing, dmPath)

    End If
    
    Music_Play = True
    Exit Function

Error_Handl:
    Music_Play = False

End Function

Public Function Music_GetLoop() As Boolean
    
    On Error GoTo Error_Handl
    
    If IMPos Is Nothing Then Exit Function
    If IMPos.StopTime = IMPos.CurrentPosition Then
        Music_Stop
        Music_Play
        Music_GetLoop = True

    End If
    
    Exit Function
    
Error_Handl:
    Music_GetLoop = False

End Function

Public Function Music_Stop() As Boolean

    On Error GoTo Error_Handl
    
    If Not IMC Is Nothing Then
        IMC.Stop
        IMPos.CurrentPosition = 0
    
        Music_Stop = True

    End If
        
    Exit Function

Error_Handl:
    Music_Stop = False

End Function

Public Function Music_Pause() As Boolean

    On Error GoTo Error_Handl
    
    IMC.Pause
    
    Music_Pause = True
    Exit Function

Error_Handl:
    Music_Pause = False

End Function

Public Function Music_Volume_Set(ByVal sound_volume As Long) As Boolean

    On Error GoTo Error_Handl
    
    If (Musica = CONST_DESHABILITADA) Then Exit Function
    
    If sound_volume > 0 Then sound_volume = 0
    If sound_volume < -4000 Then sound_volume = -4000
    
    If lastMid = False Then
        If IBA Is Nothing Then Exit Function
        IBA.Volume = sound_volume
    Else
        dmPerf.SetMasterVolume (sound_volume)

    End If

    Music_Volume_Set = True
    Exit Function

Error_Handl:
    Music_Volume_Set = False

End Function

Public Function Ambient_Load(ByVal file_name As Integer, Optional ByVal Volume As Long = 0)
    
    On Error GoTo Error_Handl
    
    If DS Is Nothing Then Exit Function
        
    Dim dsbd As DSBUFFERDESC

    dsbd.lFlags = DSBCAPS_CTRLVOLUME
    dsbd.fxFormat.nFormatTag = WAVE_FORMAT_PCM
    dsbd.fxFormat.nChannels = 2
    dsbd.fxFormat.lSamplesPerSec = 22050
    dsbd.fxFormat.nBitsPerSample = 16
    dsbd.fxFormat.nBlockAlign = dsbd.fxFormat.nBitsPerSample / 8 * dsbd.fxFormat.nChannels
    dsbd.fxFormat.lAvgBytesPerSec = dsbd.fxFormat.lSamplesPerSec * dsbd.fxFormat.nBlockAlign
    
    #If Compresion = 1 Then

        If Extract_File(wav, App.Path & "\..\Recursos\OUTPUT\", file_name & ".wav", Windows_Temp_Dir, False) Then
            Set ambient_buffer = DS.CreateSoundBufferFromFile(Windows_Temp_Dir & file_name & ".wav", dsbd)
            'Delete_File Windows_Temp_Dir & file_name & ".wav"
        Else
            GoTo Error_Handl

        End If

    #Else

        Dim PathToFile As String

        PathToFile = App.Path & "/../Recursos/WAV/" & file_name & ".wav"
    
        If FileExist(PathToFile, vbArchive) Then
            Set ambient_buffer = DS.CreateSoundBufferFromFile(PathToFile, dsbd)
        Else
            GoTo Error_Handl

        End If

    #End If

    If Volume < -4000 Then Volume = -4000
    ambient_buffer.SetVolume Volume
    
    Ambient_Load = True
    Exit Function

Error_Handl:
    Ambient_Load = False

End Function

Public Function Ambient_Play() As Boolean
    
    On Error GoTo Error_Handl
    
    If Not ambient_buffer Is Nothing Then
        ambient_buffer.Play DSBPLAY_LOOPING
        Ambient_Play = True

    End If

    Exit Function

Error_Handl:
    Ambient_Play = False

End Function

Public Function Ambient_Stop() As Boolean

    On Error GoTo Error_Handl
    
    If ambient_buffer Is Nothing Then Exit Function
    ambient_buffer.Stop
        
    Ambient_Stop = True
    Exit Function

Error_Handl:
    Ambient_Stop = False

End Function

Public Function Ambient_Volume_Set(ByVal sound_volume As Long) As Boolean

    On Error GoTo Error_Handl
    
    If ambient_buffer Is Nothing Then Exit Function
    
    If sound_volume > 0 Then sound_volume = 0
    If sound_volume < -4000 Then sound_volume = -4000
    
    ambient_buffer.SetVolume sound_volume

    Ambient_Volume_Set = True
    Exit Function

Error_Handl:
    Ambient_Volume_Set = False

End Function

Public Function Sound_Render()
    '**************************************************************
    'Author: Augusto Jos� Rando
    'Last Modify Date: 17/08/2005
    '
    '**************************************************************

    On Error Resume Next

    Dim TiempoActual    As Long, x As Integer, y As Integer

    Static under_roof   As Boolean

    Static Sound_Buffer As DirectSoundSecondaryBuffer8

    Static play_fire    As Boolean

    If UserMap > 0 Then
        under_roof = bTecho
        
        If bRain And MapDat.LLUVIA Then
            
            If bTecho Then
                If mAmbienteActual <> SND_LLUVIAIN Then
                    If (mAmbienteActual <> SND_LLUVIAIN) And (mAmbienteActual <> SND_LLUVIAOUT) Then
                        LastAmbienteActual = mAmbienteActual

                    End If
                                            
                    Call Ambient_Stop
                    mEffectsVolumeAmbient = INTERIOR_VOLUME
                    mAmbienteActual = SND_LLUVIAIN
                        
                    Call Ambient_Load(SND_LLUVIAIN, VolAmbient)
                    Call Ambient_Play

                End If

            Else

                If mAmbienteActual <> SND_LLUVIAOUT Then
                        
                    If (mAmbienteActual <> SND_LLUVIAIN) And (mAmbienteActual <> SND_LLUVIAOUT) Then
                        LastAmbienteActual = mAmbienteActual

                    End If

                    Call Ambient_Stop
                    mEffectsVolumeAmbient = EXTERIOR_VOLUME
                    mAmbienteActual = SND_LLUVIAOUT
                        
                    Call Ambient_Load(SND_LLUVIAOUT, VolAmbient)
                    Call Ambient_Play

                End If

            End If

        Else 'No aplica efectos meteorol�gicos a la zona

            If LastAmbienteActual > 0 Then
                Call Ambient_Stop
                mAmbienteActual = LastAmbienteActual
                Call Ambient_Load(mAmbienteActual, VolAmbient)
                Call Ambient_Play
                LastAmbienteActual = 0

            End If
                
            If bTecho Then 'En los interiores el sonido es m�s suave :)
                
                If mEffectsVolumeAmbient <> VolAmbient - 1000 Then

                    Call Ambient_Volume_Set(VolAmbient - 1000)
                    mEffectsVolumeAmbient = VolAmbient - 1000

                End If

            Else

                If mEffectsVolumeAmbient <> VolAmbient Then
                    Call Ambient_Volume_Set(VolAmbient)
                    mEffectsVolumeAmbient = VolAmbient

                End If
                                        
                If fX Then

                    Dim location As Position

                    If bFogata Then
                        bFogata = HayFogata(location)
                        play_fire = Sound_Play(SND_FUEGO, True, Calculate_Volume(location.x, location.y), Calculate_Pan(location.x, location.y))

                        If Moviendose Then
                            Call Buffers.CambiarVolumen(SND_FUEGO, Calculate_Volume(location.x, location.y))
                            Call Buffers.CambiarPan(SND_FUEGO, Calculate_Pan(location.x, location.y))

                        End If

                    Else
                        bFogata = HayFogata(location)
                        play_fire = Not Sound_Stop(SND_FUEGO)

                    End If

                End If

            End If

        End If

    End If

    If mFadingMusicMod <> 0 And Musica <> CONST_DESHABILITADA Then
 
        TiempoActual = (GetTickCount() And &H7FFFFFFF)
        
        '�Fade up or fade down? Primero fade down, despu�s fade up
        If TimeElapsed + 60 < TiempoActual Then
            
            TimeElapsed = TiempoActual
            
            'Fade down
            If mMusicVolume > mMusicVolumeMax Then
                
                mMusicVolume = mMusicVolumeMax
                
                If mFadingStatus = 0 Then
                    mFadingStatus = 1
                    mFadingMusicMod = mFadingMusicMod * (-1)
                Else
                    'TERMINAMOS!!!
                    mFadingMusicMod = 0
                    mFadingStatus = 0

                End If

                'Fade up
            ElseIf mMusicVolume < -4000 Then
                mFadingStatus = 2
                mMusicVolume = -4000
                mFadingMusicMod = mFadingMusicMod * (-1)

            End If
            
            mMusicVolume = mMusicVolume + mFadingMusicMod
            
            If mNextMusic > 0 And mFadingStatus = 2 Then

                If Music_Load(mNextMusic, mMusicVolume) Then
                    Music_Stop
                    Music_Play
                                
                End If

                mNextMusic = 0
            Else
                Music_Volume_Set mMusicVolume

            End If

        End If

    Else
        mMusicVolume = mMusicVolumeMax

    End If

    If Musica <> CONST_DESHABILITADA Then Call Music_GetLoop
    'CurrentUser.moved = False

End Function

Public Function Sound_Load(ByVal file_num As Integer) As Boolean
    '**************************************************************
    'Author: Augusto Jos� Rando
    'Last Modify Date: 7/16/2005
    '**************************************************************
    
    If Buffers.CargarBuffer(file_num) Then
        Sound_Load = True

    End If

End Function

Public Function Calculate_Pan(ByVal map_x As Integer, ByVal map_y As Integer) As Long
    '**************************************************************
    'Author: Augusto Jos� Rando
    'Last Modify Date: 7/16/2005
    '**************************************************************
    
    Dim total_distance As Integer, position_sgn As Integer, curr_x As Integer, curr_y As Integer

    total_distance = General_Distance_Get(map_x, map_y, UserPos.x, UserPos.y)
        
    If mInvert_Sound = False Then
        If map_x < UserPos.x Then
            position_sgn = -1
        Else
            position_sgn = 1

        End If

    Else

        If map_x > UserPos.x Then
            position_sgn = -1
        Else
            position_sgn = 1

        End If

    End If
        
    If (total_distance = 0) Or (map_x = UserPos.x) Then
        Calculate_Pan = 0
    ElseIf total_distance < 19 Then
        Calculate_Pan = position_sgn * (total_distance * 500)
    Else
        Calculate_Pan = position_sgn * 9000

    End If

End Function

Public Function Calculate_Volume(ByVal map_x As Integer, ByVal map_y As Integer) As Long
    '**************************************************************
    'Author: Augusto Jos� Rando
    'Last Modify Date: 7/16/2005
    '**************************************************************
    
    Dim total_distance As Integer, curr_x As Integer, curr_y As Integer

    total_distance = General_Distance_Get(map_x, map_y, UserPos.x, UserPos.y)
        
    If (total_distance = 0) Then
        Calculate_Volume = mEffectsVolumeMax
            
    ElseIf total_distance < 19 Then
        Calculate_Volume = mEffectsVolumeMax - (total_distance * 120)
    Else
        Calculate_Volume = mEffectsVolumeMax - 450

    End If

    If total_distance > 19 Then Calculate_Volume = -4000
    If Calculate_Volume < -4000 Then Calculate_Volume = -4000

End Function

Public Sub BorraTimer()

    If Buffers Is Nothing Then Exit Sub
    Buffers.BorraTimerProc

End Sub
