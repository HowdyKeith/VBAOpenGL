Attribute VB_Name = "EngineDemo"
Option Explicit

' ==========================================================
' CONFIGURATION
' ==========================================================
Private Const SKYBOX_IMAGE As String = "C:\downloads\map.png"
Private Const SHAPE_TEXTURE As String = "C:\downloads\map.png"
Private Const HOTSPOT_IMAGE As String = "C:\downloads\pet.jpg"
' ==========================================================

Sub RunEngine()
    Dim glWin As New OpenGLWindow
    Dim freq As Currency, startTime As Currency, currentTime As Currency
    Dim delta As Single, frameCount As Long, lastFPSUpdate As Currency

    If Not glWin.Create(1024, 768, "VBA Engine v1.42") Then Exit Sub

    glWin.InitDemo
    glWin.SetPetPath HOTSPOT_IMAGE
    
    ' Load textures
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
            glWin.UpdatePhysics
            
            frameCount = frameCount + 1
            If (currentTime - lastFPSUpdate) / freq > 0.5 Then
                glWin.UpdateFPSOverlay "FPS: " & format(frameCount / ((currentTime - lastFPSUpdate) / freq), "0.0")
                frameCount = 0: lastFPSUpdate = currentTime
            End If
            
            glWin.DrawScene
            glWin.DoSwapBuffers
        End If
        
        DoEvents
    Loop

    glWin.DrainMessages
    glWin.CleanupDemo
    glWin.Destroy
    Set glWin = Nothing
End Sub
