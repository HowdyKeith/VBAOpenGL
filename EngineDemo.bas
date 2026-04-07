Attribute VB_Name = "EngineDemo"
Option Explicit
' =============================================
' ModDemo.bas
' Version: 1.51 Physics
' =============================================

Private Const SKYBOX_IMAGE As String = "C:\downloads\map.png"
Private Const SHAPE_TEXTURE As String = "C:\downloads\map.png"
Private Const HOTSPOT_IMAGE As String = "C:\downloads\pet.jpg"

Sub RunEngine()
    Dim glWin As New OpenGLWindow
    Dim freq As Currency, startTime As Currency, currentTime As Currency
    Dim delta As Single, frameCount As Long, lastFPSUpdate As Currency
    Dim lastSheetUpdate As Double

    If Not glWin.Create(1024, 768, "VBA PhysX Engine v1.51") Then Exit Sub

    glWin.InitDemo
    glWin.SetPetPath HOTSPOT_IMAGE
    glWin.Load360Texture SKYBOX_IMAGE
    glWin.SetCubeTexture glWin.LoadTextureFromFile(SHAPE_TEXTURE)

    QueryPerformanceFrequency freq
    QueryPerformanceCounter startTime
    lastFPSUpdate = startTime

    Do While glWin.IsRunning
        QueryPerformanceCounter currentTime
        delta = (currentTime - startTime) / freq
        startTime = currentTime

        glWin.ProcessMessages

        If glWin.IsRunning Then
            If Timer - lastSheetUpdate > 0.2 Then
                On Error Resume Next
                glWin.SyncWithSpreadsheet Range("DataBuffer").Value
                On Error GoTo 0
                lastSheetUpdate = Timer
            End If

            glWin.UpdatePhysics

            frameCount = frameCount + 1
            If (currentTime - lastFPSUpdate) / freq > 0.5 Then
                glWin.UpdateFPSOverlay "FPS: " & format(frameCount / ((currentTime - lastFPSUpdate) / freq), "0.0")
                frameCount = 0
                lastFPSUpdate = currentTime
            End If

            glWin.DrawScene
            glWin.DoSwapBuffers
        End If
    Loop

    glWin.DrainMessages
    glWin.CleanupDemo
    glWin.Destroy
    Set glWin = Nothing
End Sub

Sub OnPacketReceived()
    If Not g_OpenGLWindow Is Nothing Then
        g_OpenGLWindow.TriggerDataPulse 0.5
    End If
End Sub
