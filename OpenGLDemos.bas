Attribute VB_Name = "OpenGLDemos"
Option Explicit

' ============================================================
' VBA OpenGL Demo Module - Various OpenGL Capabilities
' Demonstrates different OpenGL features available in VBA
' ============================================================

#If VBA7 Then
    ' Declarations for VBA7 (supports both 32-bit and 64-bit)
    Private Declare PtrSafe Function CreateWindowEx Lib "user32" Alias "CreateWindowExW" _
        (ByVal dwExStyle As Long, ByVal lpClassName As LongPtr, ByVal lpWindowName As LongPtr, _
         ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, _
         ByVal hWndParent As LongPtr, ByVal hMenu As LongPtr, ByVal hInstance As LongPtr, ByVal lpParam As LongPtr) As LongPtr
    Private Declare PtrSafe Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleW" (ByVal lpModuleName As LongPtr) As LongPtr
    Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
    Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hWnd As LongPtr, ByVal hdc As LongPtr) As Long
    Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal nCmdShow As Long) As Long
    Private Declare PtrSafe Function UpdateWindow Lib "user32" (ByVal hWnd As LongPtr) As Long
    Private Declare PtrSafe Function DestroyWindow Lib "user32" (ByVal hWnd As LongPtr) As Long
    Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
    Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
    Private Declare PtrSafe Function SwapBuffers Lib "gdi32" (ByVal hdc As LongPtr) As Long
    Private Declare PtrSafe Function ChoosePixelFormat Lib "gdi32" (ByVal hdc As LongPtr, ByRef pfd As PIXELFORMATDESCRIPTOR) As Long
    Private Declare PtrSafe Function SetPixelFormat Lib "gdi32" (ByVal hdc As LongPtr, ByVal format As Long, ByRef pfd As PIXELFORMATDESCRIPTOR) As Long
    Private Declare PtrSafe Function wglCreateContext Lib "opengl32.dll" (ByVal hdc As LongPtr) As LongPtr
    Private Declare PtrSafe Function wglMakeCurrent Lib "opengl32.dll" (ByVal hdc As LongPtr, ByVal hGLRC As LongPtr) As Long
    Private Declare PtrSafe Function wglDeleteContext Lib "opengl32.dll" (ByVal hGLRC As LongPtr) As Long
    Private Declare PtrSafe Sub glClearColor Lib "opengl32.dll" (ByVal r As Single, ByVal g As Single, ByVal b As Single, ByVal a As Single)
    Private Declare PtrSafe Sub glClear Lib "opengl32.dll" (ByVal mask As Long)
    Private Declare PtrSafe Sub glBegin Lib "opengl32.dll" (ByVal mode As Long)
    Private Declare PtrSafe Sub glEnd Lib "opengl32.dll" ()
    Private Declare PtrSafe Sub glVertex2f Lib "opengl32.dll" (ByVal x As Single, ByVal y As Single)
    Private Declare PtrSafe Sub glVertex3f Lib "opengl32.dll" (ByVal x As Single, ByVal y As Single, ByVal z As Single)
    Private Declare PtrSafe Sub glColor3f Lib "opengl32.dll" (ByVal r As Single, ByVal g As Single, ByVal b As Single)
    Private Declare PtrSafe Sub glColor4f Lib "opengl32.dll" (ByVal r As Single, ByVal g As Single, ByVal b As Single, ByVal a As Single)
    Private Declare PtrSafe Sub glViewport Lib "opengl32.dll" (ByVal x As Long, ByVal y As Long, ByVal width As Long, ByVal height As Long)
    Private Declare PtrSafe Sub glOrtho Lib "opengl32.dll" (ByVal Left As Double, ByVal Right As Double, ByVal Bottom As Double, ByVal Top As Double, ByVal zNear As Double, ByVal zFar As Double)
    Public Declare PtrSafe Sub glMatrixMode Lib "opengl32.dll" (ByVal mode As Long)
    Public Declare PtrSafe Sub glLoadIdentity Lib "opengl32.dll" ()
    Private Declare PtrSafe Sub glRotatef Lib "opengl32.dll" (ByVal angle As Single, ByVal x As Single, ByVal y As Single, ByVal z As Single)
    Private Declare PtrSafe Sub glTranslatef Lib "opengl32.dll" (ByVal x As Single, ByVal y As Single, ByVal z As Single)
    Private Declare PtrSafe Sub glScalef Lib "opengl32.dll" (ByVal x As Single, ByVal y As Single, ByVal z As Single)
    Private Declare PtrSafe Sub glPushMatrix Lib "opengl32.dll" ()
    Private Declare PtrSafe Sub glPopMatrix Lib "opengl32.dll" ()
    Private Declare PtrSafe Sub glEnable Lib "opengl32.dll" (ByVal cap As Long)
    Private Declare PtrSafe Sub glDisable Lib "opengl32.dll" (ByVal cap As Long)
    Private Declare PtrSafe Sub glBlendFunc Lib "opengl32.dll" (ByVal sfactor As Long, ByVal dfactor As Long)
    Private Declare PtrSafe Sub glLineWidth Lib "opengl32.dll" (ByVal width As Single)
    Private Declare PtrSafe Sub glPointSize Lib "opengl32.dll" (ByVal size As Single)
    Private Declare PtrSafe Sub glFrustum Lib "opengl32.dll" (ByVal Left As Double, ByVal Right As Double, ByVal Bottom As Double, ByVal Top As Double, ByVal zNear As Double, ByVal zFar As Double)
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
    
    Private g_hWnd As LongPtr
    Private g_hDC As LongPtr
    Private g_hGLRC As LongPtr
#Else
    Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExW" _
        (ByVal dwExStyle As Long, ByVal lpClassName As Long, ByVal lpWindowName As Long, _
         ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, _
         ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByVal lpParam As Long) As Long
    Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleW" (ByVal lpModuleName As Long) As Long
    Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
    Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
    Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
    Private Declare Function UpdateWindow Lib "user32" (ByVal hWnd As Long) As Long
    Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
    Private Declare Function GetTickCount Lib "kernel32" () As Long
    Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
    Private Declare Function SwapBuffers Lib "gdi32" (ByVal hDC As Long) As Long
    Private Declare Function ChoosePixelFormat Lib "gdi32" (ByVal hDC As Long, ByRef pfd As PIXELFORMATDESCRIPTOR) As Long
    Private Declare Function SetPixelFormat Lib "gdi32" (ByVal hDC As Long, ByVal format As Long, ByRef pfd As PIXELFORMATDESCRIPTOR) As Long
    Private Declare Function wglCreateContext Lib "opengl32.dll" (ByVal hDC As Long) As Long
    Private Declare Function wglMakeCurrent Lib "opengl32.dll" (ByVal hDC As Long, ByVal hGLRC As Long) As Long
    Private Declare Function wglDeleteContext Lib "opengl32.dll" (ByVal hGLRC As Long) As Long
    Private Declare Sub glClearColor Lib "opengl32.dll" (ByVal r As Single, ByVal g As Single, ByVal b As Single, ByVal a As Single)
    Private Declare Sub glClear Lib "opengl32.dll" (ByVal mask As Long)
    Private Declare Sub glBegin Lib "opengl32.dll" (ByVal mode As Long)
    Private Declare Sub glEnd Lib "opengl32.dll" ()
    Private Declare Sub glVertex2f Lib "opengl32.dll" (ByVal x As Single, ByVal y As Single)
    Private Declare Sub glVertex3f Lib "opengl32.dll" (ByVal x As Single, ByVal y As Single, ByVal z As Single)
    Private Declare Sub glColor3f Lib "opengl32.dll" (ByVal r As Single, ByVal g As Single, ByVal b As Single)
    Private Declare Sub glColor4f Lib "opengl32.dll" (ByVal r As Single, ByVal g As Single, ByVal b As Single, ByVal a As Single)
    Private Declare Sub glViewport Lib "opengl32.dll" (ByVal x As Long, ByVal y As Long, ByVal width As Long, ByVal height As Long)
    Private Declare Sub glOrtho Lib "opengl32.dll" (ByVal left As Double, ByVal right As Double, ByVal bottom As Double, ByVal top As Double, ByVal zNear As Double, ByVal zFar As Double)
    Private Declare Sub glMatrixMode Lib "opengl32.dll" (ByVal mode As Long)
    Private Declare Sub glLoadIdentity Lib "opengl32.dll" ()
    Private Declare Sub glRotatef Lib "opengl32.dll" (ByVal angle As Single, ByVal x As Single, ByVal y As Single, ByVal z As Single)
    Private Declare Sub glTranslatef Lib "opengl32.dll" (ByVal x As Single, ByVal y As Single, ByVal z As Single)
    Private Declare Sub glScalef Lib "opengl32.dll" (ByVal x As Single, ByVal y As Single, ByVal z As Single)
    Private Declare Sub glPushMatrix Lib "opengl32.dll" ()
    Private Declare Sub glPopMatrix Lib "opengl32.dll" ()
    Private Declare Sub glEnable Lib "opengl32.dll" (ByVal cap As Long)
    Private Declare Sub glDisable Lib "opengl32.dll" (ByVal cap As Long)
    Private Declare Sub glBlendFunc Lib "opengl32.dll" (ByVal sfactor As Long, ByVal dfactor As Long)
    Private Declare Sub glLineWidth Lib "opengl32.dll" (ByVal width As Single)
    Private Declare Sub glPointSize Lib "opengl32.dll" (ByVal size As Single)
    Private Declare Sub glFrustum Lib "opengl32.dll" (ByVal left As Double, ByVal right As Double, ByVal bottom As Double, ByVal top As Double, ByVal zNear As Double, ByVal zFar As Double)
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    
    Private g_hWnd As Long
    Private g_hDC As Long
    Private g_hGLRC As Long
#End If

Private g_Running As Boolean

' Structures
Private Type PIXELFORMATDESCRIPTOR
    nSize As Integer
    nVersion As Integer
    dwFlags As Long
    iPixelType As Byte
    cColorBits As Byte
    cRedBits As Byte
    cRedShift As Byte
    cGreenBits As Byte
    cGreenShift As Byte
    cBlueBits As Byte
    cBlueShift As Byte
    cAlphaBits As Byte
    cAlphaShift As Byte
    cAccumBits As Byte
    cAccumRedBits As Byte
    cAccumGreenBits As Byte
    cAccumBlueBits As Byte
    cAccumAlphaBits As Byte
    cDepthBits As Byte
    cStencilBits As Byte
    cAuxBuffers As Byte
    iLayerType As Byte
    bReserved As Byte
    dwLayerMask As Long
    dwVisibleMask As Long
    dwDamageMask As Long
End Type

' Constants
Private Const WS_OVERLAPPEDWINDOW As Long = &HCF0000
Private Const SW_SHOW As Long = 5
Private Const VK_ESCAPE As Long = 27
Private Const PFD_DRAW_TO_WINDOW = &H4
Private Const PFD_SUPPORT_OPENGL = &H20
Private Const PFD_DOUBLEBUFFER = &H1
Private Const PFD_TYPE_RGBA = 0
Private Const GL_COLOR_BUFFER_BIT = &H4000
Private Const GL_DEPTH_BUFFER_BIT = &H100
Private Const GL_POINTS = &H0
Private Const GL_LINES = &H1
Private Const GL_LINE_LOOP = &H2
Private Const GL_LINE_STRIP = &H3
Private Const GL_TRIANGLES = &H4
Private Const GL_QUADS = &H7
Public Const GL_PROJECTION = &H1701
Private Const GL_MODELVIEW = &H1700
Private Const GL_BLEND = &HBE2
Private Const GL_SRC_ALPHA = &H302
Private Const GL_ONE_MINUS_SRC_ALPHA = &H303
Private Const GL_DEPTH_TEST = &HB71

Sub WaitMilliseconds(ms As Long)
    Sleep ms
End Sub

' ============================================================
' Demo 1: Array Similarity Visualization
' ============================================================
Public Sub DemoArraySimilarity()
    Dim array1(1 To 10, 1 To 10) As Double
    Dim array2(1 To 10, 1 To 10) As Double
    Dim i As Long, j As Long
    
    Randomize
    For i = 1 To 10
        For j = 1 To 10
            array1(i, j) = Int(Rnd() * 100)
            If Rnd() > 0.7 Then
                array2(i, j) = Int(Rnd() * 100)
            Else
                array2(i, j) = array1(i, j) + (Rnd() - 0.5) * 20
            End If
        Next j
    Next i
    
    If InitializeOpenGL("Array Similarity Visualization") Then
        SetupOrthographic2D
        VisualizeArrayComparison array1, array2
        CleanupOpenGL
    Else
        MsgBox "Failed to initialize OpenGL for Array Similarity Visualization", vbCritical
    End If
End Sub

Private Sub VisualizeArrayComparison(arr1() As Double, arr2() As Double)
    Dim startTime As Long
    Dim differences(1 To 10, 1 To 10) As Double
    Dim maxDiff As Double, minDiff As Double
    Dim i As Long, j As Long, cellSize As Single
    
    maxDiff = -999999: minDiff = 999999
    For i = 1 To 10
        For j = 1 To 10
            differences(i, j) = Abs(arr1(i, j) - arr2(i, j))
            If differences(i, j) > maxDiff Then maxDiff = differences(i, j)
            If differences(i, j) < minDiff Then minDiff = differences(i, j)
        Next j
    Next i
    
    cellSize = 30
    startTime = GetTickCount()
    
    Do While GetTickCount() - startTime < 10000 And Not (GetAsyncKeyState(VK_ESCAPE) And &H8000)
        glClear GL_COLOR_BUFFER_BIT
        For i = 1 To 10
            For j = 1 To 10
                Dim normalizedDiff As Single
                If maxDiff = minDiff Then
                    normalizedDiff = 0
                Else
                    normalizedDiff = (differences(i, j) - minDiff) / (maxDiff - minDiff)
                End If
                glColor3f normalizedDiff, 1 - normalizedDiff, 0
                glBegin GL_QUADS
                    glVertex2f (i - 1) * cellSize + 50, (j - 1) * cellSize + 50
                    glVertex2f i * cellSize + 50, (j - 1) * cellSize + 50
                    glVertex2f i * cellSize + 50, j * cellSize + 50
                    glVertex2f (i - 1) * cellSize + 50, j * cellSize + 50
                glEnd
            Next j
        Next i
        glColor3f 0.3, 0.3, 0.3
        glBegin GL_LINES
        For i = 0 To 10
            glVertex2f 50 + i * cellSize, 50
            glVertex2f 50 + i * cellSize, 50 + 10 * cellSize
            glVertex2f 50, 50 + i * cellSize
            glVertex2f 50 + 10 * cellSize, 50 + i * cellSize
        Next i
        glEnd
        SwapBuffers g_hDC
        DoEvents
        WaitMilliseconds 50
    Loop
End Sub

' ============================================================
' Demo 2: 3D Rotating Cube
' ============================================================
Public Sub Demo3DRotatingCube()
    If InitializeOpenGL("3D Rotating Cube Demo") Then
        Setup3DPerspective
        Render3DScene
        CleanupOpenGL
    Else
        MsgBox "Failed to initialize OpenGL for 3D Rotating Cube Demo", vbCritical
    End If
End Sub

Private Sub Setup3DPerspective()
    glMatrixMode GL_PROJECTION
    glLoadIdentity
    glFrustum -1, 1, -1, 1, 2, 50
    glMatrixMode GL_MODELVIEW
    glEnable GL_DEPTH_TEST
    glClearColor 0.1, 0.1, 0.2, 1#
End Sub

Private Sub Render3DScene()
    Dim startTime As Long, angle As Single
    startTime = GetTickCount()
    
    Do While GetTickCount() - startTime < 15000 And Not (GetAsyncKeyState(VK_ESCAPE) And &H8000)
        angle = (GetTickCount() - startTime) / 20
        glClear GL_COLOR_BUFFER_BIT Or GL_DEPTH_BUFFER_BIT
        glLoadIdentity
        glTranslatef 0, 0, -8
        glRotatef angle, 1, 1, 1
        DrawColoredCube
        SwapBuffers g_hDC
        DoEvents
        WaitMilliseconds 10
    Loop
End Sub

Private Sub DrawColoredCube()
    glBegin GL_QUADS
    glColor3f 1, 0, 0
    glVertex3f -1, -1, 1: glVertex3f 1, -1, 1: glVertex3f 1, 1, 1: glVertex3f -1, 1, 1
    glColor3f 0, 1, 0
    glVertex3f -1, -1, -1: glVertex3f -1, 1, -1: glVertex3f 1, 1, -1: glVertex3f 1, -1, -1
    glColor3f 0, 0, 1
    glVertex3f -1, 1, -1: glVertex3f -1, 1, 1: glVertex3f 1, 1, 1: glVertex3f 1, 1, -1
    glColor3f 1, 1, 0
    glVertex3f -1, -1, -1: glVertex3f 1, -1, -1: glVertex3f 1, -1, 1: glVertex3f -1, -1, 1
    glColor3f 1, 0, 1
    glVertex3f 1, -1, -1: glVertex3f 1, 1, -1: glVertex3f 1, 1, 1: glVertex3f 1, -1, 1
    glColor3f 0, 1, 1
    glVertex3f -1, -1, -1: glVertex3f -1, -1, 1: glVertex3f -1, 1, 1: glVertex3f -1, 1, -1
    glEnd
End Sub

' ============================================================
' Demo 3: Simple 2D "Game" - Moving Rectangle
' ============================================================
Public Sub DemoSimple2DGame()
    If InitializeOpenGL("Simple 2D Game Demo") Then
        SetupOrthographic2D
        Run2DGameLoop
        CleanupOpenGL
    Else
        MsgBox "Failed to initialize OpenGL for Simple 2D Game Demo", vbCritical
    End If
End Sub

Private Sub Run2DGameLoop()
    Dim playerX As Single, playerY As Single
    Dim startTime As Long, Speed As Single
    Dim obstacles(1 To 5) As Single
    Dim i As Long
    
    playerX = 100: playerY = 300: Speed = 3
    For i = 1 To 5
        obstacles(i) = 800 + i * 150
    Next i
    
    startTime = GetTickCount()
    
    Do While GetTickCount() - startTime < 30000 And Not (GetAsyncKeyState(VK_ESCAPE) And &H8000)
        If GetAsyncKeyState(38) And &H8000 Then playerY = playerY - Speed ' Up
        If GetAsyncKeyState(40) And &H8000 Then playerY = playerY + Speed ' Down
        If GetAsyncKeyState(37) And &H8000 Then playerX = playerX - Speed ' Left
        If GetAsyncKeyState(39) And &H8000 Then playerX = playerX + Speed ' Right
        
        If playerX < 0 Then playerX = 0
        If playerX > 750 Then playerX = 750
        If playerY < 0 Then playerY = 0
        If playerY > 550 Then playerY = 550
        
        For i = 1 To 5
            obstacles(i) = obstacles(i) - Speed * 2
            If obstacles(i) < -50 Then obstacles(i) = 800 + Rnd() * 200
        Next i
        
        glClear GL_COLOR_BUFFER_BIT
        glColor3f 0.2, 0.5, 1
        glBegin GL_QUADS
            glVertex2f playerX, playerY
            glVertex2f playerX + 50, playerY
            glVertex2f playerX + 50, playerY + 50
            glVertex2f playerX, playerY + 50
        glEnd
        glColor3f 1, 0.2, 0.2
        For i = 1 To 5
            glBegin GL_QUADS
                glVertex2f obstacles(i), 200 + i * 50
                glVertex2f obstacles(i) + 40, 200 + i * 50
                glVertex2f obstacles(i) + 40, 240 + i * 50
                glVertex2f obstacles(i), 240 + i * 50
            glEnd
        Next i
        SwapBuffers g_hDC
        DoEvents
        WaitMilliseconds 10
    Loop
End Sub

' ============================================================
' Demo 4: Data Visualization - Real-time Graph
' ============================================================
Public Sub DemoDataVisualization()
    If InitializeOpenGL("Real-time Data Visualization") Then
        SetupOrthographic2D
        RenderDataGraph
        CleanupOpenGL
    Else
        MsgBox "Failed to initialize OpenGL for Real-time Data Visualization", vbCritical
    End If
End Sub

Private Sub RenderDataGraph()
    Dim dataPoints(1 To 100) As Single
    Dim i As Long, startTime As Long
    Dim currentIndex As Long
    
    For i = 1 To 100
        dataPoints(i) = 300
    Next i
    
    startTime = GetTickCount()
    currentIndex = 1
    
    Do While GetTickCount() - startTime < 20000 And Not (GetAsyncKeyState(VK_ESCAPE) And &H8000)
        dataPoints(currentIndex) = 300 + Sin((GetTickCount() - startTime) / 500) * 100 + Rnd() * 50 - 25
        currentIndex = currentIndex + 1
        If currentIndex > 100 Then currentIndex = 1
        
        glClear GL_COLOR_BUFFER_BIT
        glColor3f 0.5, 0.5, 0.5
        glBegin GL_LINES
            glVertex2f 50, 300: glVertex2f 750, 300
            glVertex2f 100, 100: glVertex2f 100, 500
        glEnd
        glColor3f 0.2, 0.8, 0.3
        glBegin GL_LINE_STRIP
        For i = 1 To 100
            Dim index As Long
            index = currentIndex + i - 1
            If index > 100 Then index = index - 100
            glVertex2f 100 + (i - 1) * 6, dataPoints(index)
        Next i
        glEnd
        glColor3f 1, 0.2, 0.2
        glPointSize 8
        glBegin GL_POINTS
            glVertex2f 100 + 99 * 6, dataPoints(IIf(currentIndex = 1, 100, currentIndex - 1))
        glEnd
        SwapBuffers g_hDC
        DoEvents
        WaitMilliseconds 50
    Loop
End Sub

' ============================================================
' Demo 5: Particle System
' ============================================================
Public Sub DemoParticleSystem()
    If InitializeOpenGL("Particle System Demo") Then
        SetupOrthographic2D
        glEnable GL_BLEND
        glBlendFunc GL_SRC_ALPHA, GL_ONE_MINUS_SRC_ALPHA
        RenderParticles
        CleanupOpenGL
    Else
        MsgBox "Failed to initialize OpenGL for Particle System Demo", vbCritical
    End If
End Sub

Private Sub RenderParticles()
    Dim particles(1 To 200) As Variant
    Dim i As Long, startTime As Long
    
    For i = 1 To 200
        particles(i) = Array(400 + Rnd() * 200 - 100, 400 + Rnd() * 200 - 100, _
                             (Rnd() - 0.5) * 4, (Rnd() - 0.5) * 4, _
                             Rnd(), Rnd(), Rnd(), 1#)
    Next i
    
    startTime = GetTickCount()
    
    Do While GetTickCount() - startTime < 15000 And Not (GetAsyncKeyState(VK_ESCAPE) And &H8000)
        glClear GL_COLOR_BUFFER_BIT
        For i = 1 To 200
            particles(i)(0) = particles(i)(0) + particles(i)(2)
            particles(i)(1) = particles(i)(1) + particles(i)(3)
            particles(i)(3) = particles(i)(3) + 0.1
            particles(i)(7) = particles(i)(7) - 0.005
            If particles(i)(1) > 600 Or particles(i)(7) <= 0 Then
                particles(i)(0) = 400 + Rnd() * 200 - 100
                particles(i)(1) = 50
                particles(i)(2) = (Rnd() - 0.5) * 4
                particles(i)(3) = Rnd() * -2
                particles(i)(7) = 1#
            End If
            glColor4f particles(i)(4), particles(i)(5), particles(i)(6), particles(i)(7)
            glPointSize 4
            glBegin GL_POINTS
                glVertex2f particles(i)(0), particles(i)(1)
            glEnd
        Next i
        SwapBuffers g_hDC
        DoEvents
        WaitMilliseconds 10
    Loop
End Sub

' ============================================================
' Demo 6: Mandelbrot Fractal Visualization
' ============================================================
Public Sub DemoMandelbrot()
    If InitializeOpenGL("Mandelbrot Fractal Demo") Then
        SetupOrthographic2D
        RenderMandelbrot
        CleanupOpenGL
    Else
        MsgBox "Failed to initialize OpenGL for Mandelbrot Fractal Demo", vbCritical
    End If
End Sub

Private Sub RenderMandelbrot()
    Dim startTime As Long
    Dim zoom As Single
    Dim offsetX As Single, offsetY As Single
    Dim maxIter As Integer
    Dim x As Integer, y As Integer
    Dim iter As Integer
    Dim cx As Double, cy As Double
    Dim zx As Double, zy As Double
    Dim temp As Double
    Dim color As Single
    
    maxIter = 100
    offsetX = -0.5
    offsetY = 0
    startTime = GetTickCount()
    
    Do While GetTickCount() - startTime < 20000 And Not (GetAsyncKeyState(VK_ESCAPE) And &H8000)
        zoom = 1 + (GetTickCount() - startTime) / 5000
        glClear GL_COLOR_BUFFER_BIT
        glPointSize 1
        glBegin GL_POINTS
        For x = 0 To 799 Step 2
            For y = 0 To 599 Step 2
                cx = (x / 800# - 0.5) / zoom * 3.5 + offsetX
                cy = (y / 600# - 0.5) / zoom * 2.5 + offsetY
                zx = 0: zy = 0
                iter = 0
                Do While zx * zx + zy * zy < 4 And iter < maxIter
                    temp = zx * zx - zy * zy + cx
                    zy = 2 * zx * zy + cy
                    zx = temp
                    iter = iter + 1
                Loop
                If iter = maxIter Then
                    glColor3f 0, 0, 0
                Else
                    color = iter / maxIter
                    glColor3f color, color * 0.5, 1 - color
                End If
                glVertex2f x, y
                glVertex2f x + 1, y
                glVertex2f x, y + 1
                glVertex2f x + 1, y + 1
            Next y
        Next x
        glEnd
        SwapBuffers g_hDC
        DoEvents
        WaitMilliseconds 100
    Loop
End Sub

' ============================================================
' Demo 7: Wireframe Rotating Sphere
' ============================================================
Public Sub DemoWireframeSphere()
    If InitializeOpenGL("Wireframe Sphere Demo") Then
        Setup3DPerspective
        RenderWireframeSphere
        CleanupOpenGL
    Else
        MsgBox "Failed to initialize OpenGL for Wireframe Sphere Demo", vbCritical
    End If
End Sub

Private Sub RenderWireframeSphere()
    Dim startTime As Long, angle As Single
    Dim latitude As Integer, longitude As Integer
    Dim radius As Single
    Dim phi As Single, theta As Single
    Dim phiStep As Single, thetaStep As Single
    
    radius = 2
    phiStep = 3.14159265 / 20
    thetaStep = 2 * 3.14159265 / 40
    startTime = GetTickCount()
    
    Do While GetTickCount() - startTime < 15000 And Not (GetAsyncKeyState(VK_ESCAPE) And &H8000)
        angle = (GetTickCount() - startTime) / 20
        glClear GL_COLOR_BUFFER_BIT Or GL_DEPTH_BUFFER_BIT
        glLoadIdentity
        glTranslatef 0, 0, -8
        glRotatef angle, 0, 1, 0
        glColor3f 0.8, 0.8, 1
        glLineWidth 1
        For latitude = -10 To 10
            phi = latitude * phiStep
            glBegin GL_LINE_LOOP
            For longitude = 0 To 40
                theta = longitude * thetaStep
                glVertex3f radius * Cos(phi) * Cos(theta), radius * Sin(phi), radius * Cos(phi) * Sin(theta)
            Next longitude
            glEnd
        Next latitude
        For longitude = 0 To 20
            theta = longitude * 2 * thetaStep
            glBegin GL_LINE_STRIP
            For latitude = -10 To 10
                phi = latitude * phiStep
                glVertex3f radius * Cos(phi) * Cos(theta), radius * Sin(phi), radius * Cos(phi) * Sin(theta)
            Next latitude
            glEnd
        Next longitude
        SwapBuffers g_hDC
        DoEvents
        WaitMilliseconds 10
    Loop
End Sub

' ============================================================
' Demo 8: Rotating 2D Spiral
' ============================================================
Public Sub DemoRotatingSpiral()
    If InitializeOpenGL("Rotating Spiral Demo") Then
        SetupOrthographic2D
        RenderRotatingSpiral
        CleanupOpenGL
    Else
        MsgBox "Failed to initialize OpenGL for Rotating Spiral Demo", vbCritical
    End If
End Sub

Private Sub RenderRotatingSpiral()
    Dim startTime As Long, angle As Single
    Dim i As Long
    Dim radius As Single, theta As Single
    Dim points As Long
    
    points = 200
    startTime = GetTickCount()
    
    Do While GetTickCount() - startTime < 15000 And Not (GetAsyncKeyState(VK_ESCAPE) And &H8000)
        angle = (GetTickCount() - startTime) / 1000 * 2 * 3.14159265
        glClear GL_COLOR_BUFFER_BIT
        glColor3f 0.2, 0.8, 0.8
        glBegin GL_LINE_LOOP
        For i = 0 To points
            theta = i / 20# * 2 * 3.14159265 + angle
            radius = i / 20#
            glVertex2f 400 + radius * Cos(theta) * 50, 300 + radius * Sin(theta) * 50
        Next i
        glEnd
        SwapBuffers g_hDC
        DoEvents
        WaitMilliseconds 10
    Loop
End Sub

' ============================================================
' Demo 9: 3D Terrain Flyover
' ============================================================
Public Sub DemoTerrainFlyover()
    If InitializeOpenGL("3D Terrain Flyover Demo") Then
        Setup3DPerspective
        RenderTerrainFlyover
        CleanupOpenGL
    Else
        MsgBox "Failed to initialize OpenGL for 3D Terrain Flyover Demo", vbCritical
    End If
End Sub

Private Sub RenderTerrainFlyover()
    Dim startTime As Long
    Dim x As Integer, z As Integer
    Dim height(0 To 20, 0 To 20) As Single
    Dim i As Long, j As Long
    Dim cameraZ As Single
    
    Randomize
    For i = 0 To 20
        For j = 0 To 20
            height(i, j) = Sin(i / 5#) * Cos(j / 5#) * 2 + Rnd() * 0.5
        Next j
    Next i
    
    startTime = GetTickCount()
    
    Do While GetTickCount() - startTime < 20000 And Not (GetAsyncKeyState(VK_ESCAPE) And &H8000)
        cameraZ = (GetTickCount() - startTime) / 1000 * 2
        glClear GL_COLOR_BUFFER_BIT Or GL_DEPTH_BUFFER_BIT
        glLoadIdentity
        glTranslatef -10, -3, -15 - cameraZ
        glRotatef 30, 1, 0, 0
        glColor3f 0.3, 0.6, 0.3
        For x = 0 To 19
            For z = 0 To 19
                glBegin GL_QUADS
                    glVertex3f x, height(x, z), z
                    glVertex3f x + 1, height(x + 1, z), z
                    glVertex3f x + 1, height(x + 1, z + 1), z + 1
                    glVertex3f x, height(x, z + 1), z + 1
                glEnd
            Next z
        Next x
        SwapBuffers g_hDC
        DoEvents
        WaitMilliseconds 10
    Loop
End Sub

' ============================================================
' Demo 10: Bouncing Balls
' ============================================================
Public Sub DemoBouncingBalls()
    If InitializeOpenGL("Bouncing Balls Demo") Then
        SetupOrthographic2D
        glEnable GL_BLEND
        glBlendFunc GL_SRC_ALPHA, GL_ONE_MINUS_SRC_ALPHA
        RenderBouncingBalls
        CleanupOpenGL
    Else
        MsgBox "Failed to initialize OpenGL for Bouncing Balls Demo", vbCritical
    End If
End Sub

Private Sub RenderBouncingBalls()
    Dim balls(1 To 10) As Variant ' x, y, vx, vy, radius, r, g, b
    Dim i As Long, j As Long
    Dim startTime As Long
    Dim gravity As Single, elasticity As Single
    
    gravity = 0.2
    elasticity = 0.8
    Randomize
    For i = 1 To 10
        balls(i) = Array(100 + Rnd() * 600, 100 + Rnd() * 400, (Rnd() - 0.5) * 8, (Rnd() - 0.5) * 8, 20, Rnd(), Rnd(), Rnd())
    Next i
    
    startTime = GetTickCount()
    
    Do While GetTickCount() - startTime < 15000 And Not (GetAsyncKeyState(VK_ESCAPE) And &H8000)
        glClear GL_COLOR_BUFFER_BIT
        For i = 1 To 10
            ' Update position and velocity
            balls(i)(0) = balls(i)(0) + balls(i)(2)
            balls(i)(1) = balls(i)(1) + balls(i)(3)
            balls(i)(3) = balls(i)(3) + gravity
            
            ' Boundary collision
            If balls(i)(0) - balls(i)(4) < 0 Then
                balls(i)(0) = balls(i)(4)
                balls(i)(2) = -balls(i)(2) * elasticity
            ElseIf balls(i)(0) + balls(i)(4) > 800 Then
                balls(i)(0) = 800 - balls(i)(4)
                balls(i)(2) = -balls(i)(2) * elasticity
            End If
            If balls(i)(1) - balls(i)(4) < 0 Then
                balls(i)(1) = balls(i)(4)
                balls(i)(3) = -balls(i)(3) * elasticity
            ElseIf balls(i)(1) + balls(i)(4) > 600 Then
                balls(i)(1) = 600 - balls(i)(4)
                balls(i)(3) = -balls(i)(3) * elasticity
            End If
            
            ' Simple ball-ball collision
            For j = i + 1 To 10
                Dim dx As Single, dy As Single, dist As Single
                dx = balls(i)(0) - balls(j)(0)
                dy = balls(i)(1) - balls(j)(1)
                dist = Sqr(dx * dx + dy * dy)
                If dist < balls(i)(4) + balls(j)(4) Then
                    Dim nx As Single, ny As Single
                    nx = dx / dist
                    ny = dy / dist
                    Dim dvx As Single, dvy As Single
                    dvx = balls(i)(2) - balls(j)(2)
                    dvy = balls(i)(3) - balls(j)(3)
                    Dim impulse As Single
                    impulse = (dx * dvx + dy * dvy) / dist
                    balls(i)(2) = balls(i)(2) - impulse * nx * elasticity
                    balls(i)(3) = balls(i)(3) - impulse * ny * elasticity
                    balls(j)(2) = balls(j)(2) + impulse * nx * elasticity
                    balls(j)(3) = balls(j)(3) + impulse * ny * elasticity
                End If
            Next j
            
            ' Render ball
            glColor3f balls(i)(5), balls(i)(6), balls(i)(7)
            glBegin GL_LINE_LOOP
            Dim theta As Single, k As Integer
            For k = 0 To 20
                theta = k / 20# * 2 * 3.14159265
                glVertex2f balls(i)(0) + balls(i)(4) * Cos(theta), balls(i)(1) + balls(i)(4) * Sin(theta)
            Next k
            glEnd
        Next i
        SwapBuffers g_hDC
        DoEvents
        WaitMilliseconds 10
    Loop
End Sub

' ============================================================
' Demo 11: 3D Torus
' ============================================================
Public Sub DemoTorus()
    If InitializeOpenGL("3D Torus Demo") Then
        Setup3DPerspective
        RenderTorus
        CleanupOpenGL
    Else
        MsgBox "Failed to initialize OpenGL for 3D Torus Demo", vbCritical
    End If
End Sub

Private Sub RenderTorus()
    Dim startTime As Long, angle As Single
    Dim i As Integer, j As Integer
    Dim r As Single
    Dim theta As Single, phi As Single
    Dim thetaStep As Single, phiStep As Single
    
    r = 2 ' Major radius
    r = 0.5 ' Minor radius
    thetaStep = 2 * 3.14159265 / 30
    phiStep = 2 * 3.14159265 / 20
    startTime = GetTickCount()
    
    Do While GetTickCount() - startTime < 15000 And Not (GetAsyncKeyState(VK_ESCAPE) And &H8000)
        angle = (GetTickCount() - startTime) / 20
        glClear GL_COLOR_BUFFER_BIT Or GL_DEPTH_BUFFER_BIT
        glLoadIdentity
        glTranslatef 0, 0, -8
        glRotatef angle, 1, 1, 0
        glColor3f 0.7, 0.7, 1
        glLineWidth 1
        For i = 0 To 29
            theta = i * thetaStep
            For j = 0 To 19
                phi = j * phiStep
                glBegin GL_LINE_LOOP
                Dim k As Integer
                For k = 0 To 20
                    phi = (j + k / 20#) * phiStep
                    glVertex3f (r + r * Cos(phi)) * Cos(theta), (r + r * Cos(phi)) * Sin(theta), r * Sin(phi)
                Next k
                glEnd
            Next j
        Next i
        SwapBuffers g_hDC
        DoEvents
        WaitMilliseconds 10
    Loop
End Sub

' ============================================================
' Utility Functions
' ============================================================
Public Function InitializeOpenGL(windowTitle As String) As Boolean
    Dim pfd As PIXELFORMATDESCRIPTOR
    Dim pf As Long
    #If VBA7 Then
        Dim hInstance As LongPtr
        hInstance = GetModuleHandle(0)
        g_hWnd = CreateWindowEx(0, StrPtr("STATIC"), StrPtr(windowTitle), _
                               WS_OVERLAPPEDWINDOW, 100, 100, 800, 600, 0, 0, hInstance, 0)
    #Else
        Dim hInstance As Long
        hInstance = GetModuleHandle(0)
        g_hWnd = CreateWindowEx(0, StrPtr("STATIC"), StrPtr(windowTitle), _
                               WS_OVERLAPPEDWINDOW, 100, 100, 800, 600, 0, 0, hInstance, 0)
    #End If
    
    If g_hWnd = 0 Then
        MsgBox "Failed to create window", vbCritical
        Exit Function
    End If
    
    ShowWindow g_hWnd, SW_SHOW
    UpdateWindow g_hWnd
    
    g_hDC = GetDC(g_hWnd)
    If g_hDC = 0 Then
        MsgBox "Failed to get device context", vbCritical
        DestroyWindow g_hWnd
        g_hWnd = 0
        Exit Function
    End If
    
    With pfd
        .nSize = LenB(pfd): .nVersion = 1
        .dwFlags = PFD_DRAW_TO_WINDOW Or PFD_SUPPORT_OPENGL Or PFD_DOUBLEBUFFER
        .iPixelType = PFD_TYPE_RGBA: .cColorBits = 32: .cDepthBits = 24
    End With
    
    pf = ChoosePixelFormat(g_hDC, pfd)
    If pf = 0 Then
        MsgBox "Failed to choose pixel format", vbCritical
        ReleaseDC g_hWnd, g_hDC
        DestroyWindow g_hWnd
        g_hWnd = 0: g_hDC = 0
        Exit Function
    End If
    
    If SetPixelFormat(g_hDC, pf, pfd) = 0 Then
        MsgBox "Failed to set pixel format", vbCritical
        ReleaseDC g_hWnd, g_hDC
        DestroyWindow g_hWnd
        g_hWnd = 0: g_hDC = 0
        Exit Function
    End If
    
    g_hGLRC = wglCreateContext(g_hDC)
    If g_hGLRC = 0 Then
        MsgBox "Failed to create OpenGL context", vbCritical
        ReleaseDC g_hWnd, g_hDC
        DestroyWindow g_hWnd
        g_hWnd = 0: g_hDC = 0
        Exit Function
    End If
    
    If wglMakeCurrent(g_hDC, g_hGLRC) = 0 Then
        MsgBox "Failed to make OpenGL context current", vbCritical
        wglDeleteContext g_hGLRC
        ReleaseDC g_hWnd, g_hDC
        DestroyWindow g_hWnd
        g_hWnd = 0: g_hDC = 0: g_hGLRC = 0
        Exit Function
    End If
    
    glViewport 0, 0, 800, 600
    InitializeOpenGL = True
End Function

Private Sub SetupOrthographic2D()
    glMatrixMode GL_PROJECTION
    glLoadIdentity
    glOrtho 0, 800, 600, 0, -1, 1
    glMatrixMode GL_MODELVIEW
    glLoadIdentity
    glClearColor 0.1, 0.1, 0.1, 1#
End Sub

Public Sub CleanupOpenGL()
    If g_hGLRC <> 0 Then
        wglMakeCurrent 0, 0
        wglDeleteContext g_hGLRC
        g_hGLRC = 0
    End If
    If g_hDC <> 0 Then
        ReleaseDC g_hWnd, g_hDC
        g_hDC = 0
    End If
    If g_hWnd <> 0 Then
        DestroyWindow g_hWnd
        g_hWnd = 0
    End If
    MsgBox "Demo completed! Press any key to continue."
End Sub

' ============================================================
' Array Comparison Utility Function
' ============================================================
Public Function CompareArrays2D(arr1() As Double, arr2() As Double) As Double
    Dim i As Long, j As Long
    Dim maxRows As Long, maxCols As Long
    Dim totalElements As Long, matchingElements As Long
    Dim tolerance As Double
    
    tolerance = 0.001
    maxRows = UBound(arr1, 1)
    maxCols = UBound(arr1, 2)
    
    If maxRows <> UBound(arr2, 1) Or maxCols <> UBound(arr2, 2) Then
        CompareArrays2D = 0
        Exit Function
    End If
    
    totalElements = maxRows * maxCols
    For i = LBound(arr1, 1) To UBound(arr1, 1)
        For j = LBound(arr1, 2) To UBound(arr1, 2)
            If Abs(arr1(i, j) - arr2(i, j)) <= tolerance Then
                matchingElements = matchingElements + 1
            End If
        Next j
    Next i
    CompareArrays2D = matchingElements / totalElements
End Function

' ============================================================
' Main Demo Launcher
' ============================================================
Public Sub RunAllOpenGLDemos()
    MsgBox "Starting OpenGL Demos! Press ESC to exit any demo early.", vbInformation
    DemoArraySimilarity
    Demo3DRotatingCube
    DemoSimple2DGame
    DemoDataVisualization
    DemoParticleSystem
    DemoMandelbrot
    DemoWireframeSphere
    DemoRotatingSpiral
    DemoTerrainFlyover
    DemoBouncingBalls
    DemoTorus
    MsgBox "All demos completed!", vbInformation
End Sub

' ============================================================
' Initialize OpenGL with a given HWND (for embedding in a Form)
' ============================================================
Public Function InitializeOpenGLWithHWND(hWnd As LongPtr) As Boolean
    Dim pfd As PIXELFORMATDESCRIPTOR
    Dim pf As Long
    
    g_hWnd = hWnd
    g_hDC = GetDC(g_hWnd)
    If g_hDC = 0 Then Exit Function
    
    With pfd
        .nSize = LenB(pfd): .nVersion = 1
        .dwFlags = PFD_DRAW_TO_WINDOW Or PFD_SUPPORT_OPENGL Or PFD_DOUBLEBUFFER
        .iPixelType = PFD_TYPE_RGBA: .cColorBits = 32: .cDepthBits = 24
    End With
    
    pf = ChoosePixelFormat(g_hDC, pfd)
    If pf = 0 Then Exit Function
    If SetPixelFormat(g_hDC, pf, pfd) = 0 Then Exit Function
    
    g_hGLRC = wglCreateContext(g_hDC)
    If g_hGLRC = 0 Then Exit Function
    If wglMakeCurrent(g_hDC, g_hGLRC) = 0 Then Exit Function
    
    InitializeOpenGLWithHWND = True
End Function

' ============================================================
' Show OpenGL Demo inside a VBA UserForm
' ============================================================
Public Sub ShowDemoInForm(demoName As String)
    Dim hForm As LongPtr
    
    ' Load the form
    frmOpenGLDisplay.Show vbModeless
    DoEvents
    
    ' Get the hWnd of the form
    #If VBA7 Then
        hForm = FindWindowEx(0, 0, "ThunderDFrame", vbNullString) ' For VBA7 forms
    #Else
        hForm = FindWindowEx(0, 0, "ThunderDFrame", vbNullString) ' For 32-bit VBA forms
    #End If
    
    ' Use hForm as the parent window for OpenGL
    If InitializeOpenGLWithParent(hForm, frmOpenGLDisplay.Caption) Then
        Select Case demoName
            Case "ArraySimilarity": SetupOrthographic2D: VisualizeArrayComparisonDemo
            Case "3DRotatingCube": Setup3DPerspective: Render3DScene
            Case "Simple2DGame": SetupOrthographic2D: Run2DGameLoop
            Case "DataVisualization": SetupOrthographic2D: RenderDataGraph
            Case "ParticleSystem": SetupOrthographic2D: RenderParticles
            Case "Mandelbrot": SetupOrthographic2D: RenderMandelbrot
            Case "WireframeSphere": Setup3DPerspective: RenderWireframeSphere
            Case "RotatingSpiral": SetupOrthographic2D: RenderRotatingSpiral
            Case "TerrainFlyover": Setup3DPerspective: RenderTerrainFlyover
            Case "BouncingBalls": SetupOrthographic2D: RenderBouncingBalls
            Case "Torus": Setup3DPerspective: RenderTorus
            Case Else
                MsgBox "Unknown demo: " & demoName, vbExclamation
        End Select
        CleanupOpenGL
    Else
        MsgBox "Failed to initialize OpenGL in form", vbCritical
    End If
End Sub

Public Function InitializeOpenGLWithParent(hParent As LongPtr, windowTitle As String) As Boolean
    Dim pfd As PIXELFORMATDESCRIPTOR
    Dim pf As Long
    
    #If VBA7 Then
        g_hWnd = CreateWindowEx(0, StrPtr("STATIC"), StrPtr(windowTitle), _
                                WS_OVERLAPPEDWINDOW, 0, 0, 800, 600, hParent, 0, GetModuleHandle(0), 0)
    #Else
        g_hWnd = CreateWindowEx(0, StrPtr("STATIC"), StrPtr(windowTitle), _
                                WS_OVERLAPPEDWINDOW, 0, 0, 800, 600, hParent, 0, GetModuleHandle(0), 0)
    #End If
    
    If g_hWnd = 0 Then Exit Function
    
    ShowWindow g_hWnd, SW_SHOW
    UpdateWindow g_hWnd
    g_hDC = GetDC(g_hWnd)
    
    With pfd
        .nSize = LenB(pfd): .nVersion = 1
        .dwFlags = PFD_DRAW_TO_WINDOW Or PFD_SUPPORT_OPENGL Or PFD_DOUBLEBUFFER
        .iPixelType = PFD_TYPE_RGBA: .cColorBits = 32: .cDepthBits = 24
    End With
    
    pf = ChoosePixelFormat(g_hDC, pfd)
    If pf = 0 Then Exit Function
    If SetPixelFormat(g_hDC, pf, pfd) = 0 Then Exit Function
    
    g_hGLRC = wglCreateContext(g_hDC)
    If g_hGLRC = 0 Then Exit Function
    If wglMakeCurrent(g_hDC, g_hGLRC) = 0 Then Exit Function
    
    InitializeOpenGLWithParent = True
End Function


