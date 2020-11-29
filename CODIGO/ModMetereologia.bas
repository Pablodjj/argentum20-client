Attribute VB_Name = "ModMetereologia"
Option Explicit

Public Const LIGHT_TRANSITION_DURATION = 5000

Public Const STEP_LIGHT_TRANSITION = 1 / LIGHT_TRANSITION_DURATION

'Status
Private Const Normal        As Byte = 0
Private Const NUBLADO       As Byte = 1
Private Const LLUVIA        As Byte = 2
Private Const NIEVE         As Byte = 3
Private Const TORMENTA      As Byte = 4

Private DayColors()         As RGBA
Private DeathColor          As RGBA
Private TimeIndex           As Integer

Private NightIndex          As Integer
Private MorningIndex        As Integer

Public MeteoParticle        As Integer

Public Sub IniciarMeteorologia()
    ReDim DayColors(11)

    ' 00:00 - 02:00
    Call SetRGBA(DayColors(0), 130, 130, 130)
    NightIndex = 0
    ' 02:00 - 04:00
    Call SetRGBA(DayColors(1), 130, 130, 160)
    ' 04:00 - 06:00
    Call SetRGBA(DayColors(2), 150, 150, 180)
    ' 06:00 - 08:00
    Call SetRGBA(DayColors(3), 200, 200, 190)
    MorningIndex = 3
    ' 08:00 - 10:00
    Call SetRGBA(DayColors(4), 230, 200, 190)
    ' 10:00 - 12:00
    Call SetRGBA(DayColors(5), 255, 230, 190)
    ' 12:00 - 14:00
    Call SetRGBA(DayColors(6), 255, 240, 180)
    ' 14:00 - 16:00
    Call SetRGBA(DayColors(7), 255, 250, 170)
    ' 16:00 - 18:00
    Call SetRGBA(DayColors(8), 255, 230, 150)
    ' 18:00 - 20:00
    Call SetRGBA(DayColors(9), 255, 210, 140)
    ' 20:00 - 22:00
    Call SetRGBA(DayColors(10), 180, 150, 130)
    ' 22:00 - 00:00
    Call SetRGBA(DayColors(11), 150, 140, 130)

    ' Muerto
    Call SetRGBA(DeathColor, 120, 120, 120)
    
    TimeIndex = -1

End Sub

Public Sub RevisarHoraMundo(Optional ByVal Instantaneo As Boolean = False)

    Dim Elapsed As Single
    Elapsed = (FrameTime - HoraMundo) / DuracionDia

    Dim HoraActual As Long
    HoraActual = Fix((Elapsed - Fix(Elapsed)) * 24)
    
    Dim CurrentIndex As Integer
    CurrentIndex = HoraActual \ 2
    
    If CurrentIndex <> TimeIndex Then
        TimeIndex = CurrentIndex
        
        If Instantaneo Then
            global_light = DayColors(TimeIndex)
        Else
            Call ActualizarLuz(DayColors(TimeIndex))
        End If
        
        If TimeIndex = NightIndex Then
            Call Sound.Sound_Play(FXSound.Lobo_Sound, False, 0, 0)

        ElseIf TimeIndex = MorningIndex Then
            Call Sound.Sound_Play(FXSound.Gallo_Sound, False, 0, 0)

        End If
    End If

End Sub

Public Sub ActualizarLuz(Color As RGBA)
    last_light = global_light
    next_light = Color
    light_transition = 0#
End Sub

Public Sub RestaurarLuz()
    If UserEstado = 1 Then
        global_light = DeathColor
    Else
        global_light = DayColors(TimeIndex)
    End If
    light_transition = 1#
End Sub

Public Function EsNoche() As Boolean
    EsNoche = (TimeIndex >= NightIndex And TimeIndex < MorningIndex)
End Function
