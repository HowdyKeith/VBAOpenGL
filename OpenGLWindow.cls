VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OpenGLWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' =============================================
' OpenGLWindow.cls
' Version: 1.51 Physics
' =============================================

Private Const PI As Double = 3.14159265358979

' --- UI Constants ---
Private Const BTN_X As Long = 10
Private Const BTN_W As Long = 128
Private Const BTN_HEIGHT As Long = 35
Private Const BTN_GAP As Long = 10
Private Const UI_START_Y As Long = 50

' --- Physics Constants ---
Private Const BALL_RADIUS As Single = 1.2
Private Const ROOM_SIZE As Single = 8#
Private Const FRICTION As Single = 0.98
Private Const SPIN_FRICTION As Single = 0.995   ' spin decays slower than linear motion
Private Const SPIN_TRANSFER As Single = 0.15    ' how much collision transfers to spin

' --- Interaction & Core ---
Private m_LastX As Long
Private m_LastY As Long
Private m_IsDragging As Boolean
Public RotX As Single, RotY As Single, ZoomZ As Single
Public hWnd As LongPtr, hDC As LongPtr, hRC As LongPtr
Public IsRunning As Boolean

' --- View & Assets ---
Private m_ViewMode As Integer
Private m_ShapeMode As Integer
Private m_SphereTexture As Long
Private m_CubeTexture As Long
Private m_PetTexture As Long
Private m_SphereYaw As Double
Private m_SpherePitch As Double
Private m_SphereFOV As Double
Private token As LongPtr, font As LongPtr, brush As LongPtr
Private fpsTexture As Long, m_ResetTexture As Long
Private m_Btn360 As Long, m_BtnShape As Long
Private m_PetFilePath As String

' --- Data Visualization ---
Private m_DataBuffer(1 To 10) As Single

' --- Particles ---
Private Type Particle
    Pos(2) As Single
    Vel(2) As Single
    Life As Single
End Type
Private m_Particles(0 To 99) As Particle

' --- Physics Entities ---
' Now includes spin state per ball
Private Type PhysicalObject
    Pos(2) As Single        ' world position
    Vel(2) As Single        ' linear velocity
    SpinX As Single         ' angular velocity around X
    SpinY As Single         ' angular velocity around Y
    SpinZ As Single         ' angular velocity around Z
    AngX As Single          ' accumulated rotation angle around X
    AngY As Single          ' accumulated rotation angle around Y
    AngZ As Single          ' accumulated rotation angle around Z
End Type
Private m_Balls(0 To 4) As PhysicalObject

' --- Helper ---
Private Function GetButtonY(index As Long) As Long
    GetButtonY = UI_START_Y + (index * (BTN_HEIGHT + BTN_GAP))
End Function

Public Sub SetPetPath(ByVal path As String)
    m_PetFilePath = path
End Sub

Public Function Create(Width As Long, Height As Long, Title As String) As Boolean
    Dim wc As WNDCLASSEX
    Dim hInst As LongPtr
    hInst = GetModuleHandle(vbNullString)

    With wc
        .cbSize = LenB(wc)
        .style = CS_HREDRAW Or CS_VREDRAW Or CS_OWNDC
        .lpfnWndProc = SetWndProc(AddressOf GlobalWndProc)
        .hInstance = hInst
        .hCursor = LoadCursor(0, IDC_ARROW)
        .lpszClassName = "VBA_PhysX_Class"
    End With

    RegisterClassEx wc
    Set g_OpenGLWindow = Me

    hWnd = CreateWindowEx(0, wc.lpszClassName, Title, _
        WS_OVERLAPPEDWINDOW Or WS_CLIPCHILDREN Or WS_CLIPSIBLINGS, _
        CW_USEDEFAULT, CW_USEDEFAULT, Width, Height, 0, 0, hInst, 0)

    If hWnd <> 0 Then
        hDC = GetDC(hWnd)
        SetupPixelFormat
        hRC = wglCreateContext(hDC)
        wglMakeCurrent hDC, hRC
        ShowWindow hWnd, SW_SHOW
        UpdateWindow hWnd
        IsRunning = True
        ZoomZ = -20#
        m_ViewMode = 1
        m_SphereFOV = 80#
        Create = True
    End If
End Function

Private Sub SetupPixelFormat()
    Dim pfd As PIXELFORMATDESCRIPTOR
    pfd.nSize = LenB(pfd): pfd.nVersion = 1
    pfd.dwFlags = PFD_DRAW_TO_WINDOW Or PFD_SUPPORT_OPENGL Or PFD_DOUBLEBUFFER
    pfd.iPixelType = PFD_TYPE_RGBA: pfd.cColorBits = 32: pfd.cDepthBits = 24
    SetPixelFormat hDC, ChoosePixelFormat(hDC, pfd), pfd
End Sub

Private Sub InitDefaultDataBuffer()
    ' Nice sine wave default so bars look meaningful from the start
    Dim i As Long
    For i = 1 To 10
        m_DataBuffer(i) = 2# + 2# * Sin(i * 0.7)
    Next i
End Sub

Private Sub InitBalls()
    Dim i As Long
    For i = 0 To 4
        m_Balls(i).Pos(0) = (i - 2) * 3#
        m_Balls(i).Pos(1) = (Rnd - 0.5) * 2#
        m_Balls(i).Pos(2) = (Rnd - 0.5) * 2#
        m_Balls(i).Vel(0) = (Rnd - 0.5) * 0.15
        m_Balls(i).Vel(1) = (Rnd - 0.5) * 0.15
        m_Balls(i).Vel(2) = (Rnd - 0.5) * 0.15
        m_Balls(i).SpinX = 0
        m_Balls(i).SpinY = 0
        m_Balls(i).SpinZ = 0
        m_Balls(i).AngX = 0
        m_Balls(i).AngY = 0
        m_Balls(i).AngZ = 0
    Next i
End Sub

Public Sub InitDemo()
    Dim gpInput As GDIPlusStartupInput
    gpInput.GdiplusVersion = 1
    GdiplusStartup token, gpInput, ByVal 0&

    Dim family As LongPtr
    GdipCreateFontFamilyFromName StrPtr("Consolas"), 0, family
    GdipCreateFont family, 11, 0, 3, font
    GdipCreateSolidFill &HFFFFFFFF, brush
    GdipDeleteFontFamily family

    glGenTextures 1, fpsTexture
    glGenTextures 1, m_ResetTexture
    glGenTextures 1, m_Btn360
    glGenTextures 1, m_BtnShape

    glEnable GL_LIGHTING
    glEnable GL_LIGHT0
    glEnable GL_DEPTH_TEST
    glShadeModel GL_SMOOTH
    glEnable GL_NORMALIZE

    Dim lightPos(3) As Single
    lightPos(0) = 5#: lightPos(1) = 5#: lightPos(2) = 10#: lightPos(3) = 1#
    glLightfv GL_LIGHT0, GL_POSITION, lightPos(0)

    Dim lightAmbient(3) As Single
    lightAmbient(0) = 0.2: lightAmbient(1) = 0.2: lightAmbient(2) = 0.2: lightAmbient(3) = 1#
    glLightfv GL_LIGHT0, GL_AMBIENT, lightAmbient(0)

    Dim lightDiffuse(3) As Single
    lightDiffuse(0) = 0.8: lightDiffuse(1) = 0.8: lightDiffuse(2) = 0.8: lightDiffuse(3) = 1#
    glLightfv GL_LIGHT0, GL_DIFFUSE, lightDiffuse(0)

    Dim lightSpecular(3) As Single
    lightSpecular(0) = 1#: lightSpecular(1) = 1#: lightSpecular(2) = 1#: lightSpecular(3) = 1#
    glLightfv GL_LIGHT0, GL_SPECULAR, lightSpecular(0)

    InitDefaultDataBuffer
    InitBalls

    UpdateButtonTexture m_ResetTexture, "RESET PHYSICS", &HFF444444
    UpdateButtonTexture m_Btn360, "TOGGLE SKYBOX", &HFF664422
    UpdateButtonTexture m_BtnShape, "TOGGLE SHAPE", &HFF224466
    UpdateButtonTexture fpsTexture, "FPS: --", &H88333333
End Sub

Public Sub SetCubeTexture(ByVal texID As Long)
    m_CubeTexture = texID
End Sub

Public Sub TriggerDataPulse(burstIntensity As Single)
    Dim i As Long
    For i = 0 To 99
        m_Particles(i).Pos(0) = 0
        m_Particles(i).Pos(1) = 0
        m_Particles(i).Pos(2) = 0
        m_Particles(i).Vel(0) = (Rnd - 0.5) * burstIntensity
        m_Particles(i).Vel(1) = (Rnd - 0.5) * burstIntensity
        m_Particles(i).Vel(2) = (Rnd - 0.5) * burstIntensity
        m_Particles(i).Life = 1#
    Next i
End Sub

Public Sub SyncWithSpreadsheet(dataArray As Variant)
    Dim i As Long
    If Not IsEmpty(dataArray) Then
        For i = 1 To 10
            If i <= UBound(dataArray, 1) Then
                m_DataBuffer(i) = dataArray(i, 1)
            End If
        Next i
    End If
End Sub

Public Sub UpdateFPSOverlay(ByVal displayText As String)
    UpdateButtonTexture fpsTexture, displayText, &H88333333
End Sub

Public Sub UpdatePhysics()
    Dim i As Long, j As Long, k As Long
    Dim d As Single, dx As Single, dy As Single, dz As Single
    Dim overlap As Single, nx As Single, ny As Single, nz As Single
    Dim relVx As Single, relVy As Single, relVz As Single
    Dim relSpeed As Single
    Dim tv(2) As Single

    ' --- Particles ---
    For i = 0 To 99
        If m_Particles(i).Life > 0 Then
            For k = 0 To 2
                m_Particles(i).Pos(k) = m_Particles(i).Pos(k) + m_Particles(i).Vel(k)
            Next k
            m_Particles(i).Life = m_Particles(i).Life - 0.02
        End If
    Next i

    ' --- Balls ---
    For i = 0 To 4

        ' Apply linear friction and integrate position
        For k = 0 To 2
            m_Balls(i).Vel(k) = m_Balls(i).Vel(k) * FRICTION
            m_Balls(i).Pos(k) = m_Balls(i).Pos(k) + m_Balls(i).Vel(k)
        Next k

        ' Apply spin friction
        m_Balls(i).SpinX = m_Balls(i).SpinX * SPIN_FRICTION
        m_Balls(i).SpinY = m_Balls(i).SpinY * SPIN_FRICTION
        m_Balls(i).SpinZ = m_Balls(i).SpinZ * SPIN_FRICTION

        ' Accumulate rotation angles from spin
        m_Balls(i).AngX = m_Balls(i).AngX + m_Balls(i).SpinX
        m_Balls(i).AngY = m_Balls(i).AngY + m_Balls(i).SpinY
        m_Balls(i).AngZ = m_Balls(i).AngZ + m_Balls(i).SpinZ

        ' Wall bounce - also transfer wall normal into spin
        For k = 0 To 2
            If Abs(m_Balls(i).Pos(k)) > ROOM_SIZE Then
                ' On wall bounce, give a little spin kick perpendicular to wall
                Dim wallSpin As Single
                wallSpin = m_Balls(i).Vel(k) * SPIN_TRANSFER
                If k = 0 Then
                    m_Balls(i).SpinY = m_Balls(i).SpinY + wallSpin
                    m_Balls(i).SpinZ = m_Balls(i).SpinZ + wallSpin * 0.5
                ElseIf k = 1 Then
                    m_Balls(i).SpinX = m_Balls(i).SpinX + wallSpin
                    m_Balls(i).SpinZ = m_Balls(i).SpinZ + wallSpin * 0.5
                Else
                    m_Balls(i).SpinX = m_Balls(i).SpinX + wallSpin * 0.5
                    m_Balls(i).SpinY = m_Balls(i).SpinY + wallSpin
                End If
                m_Balls(i).Vel(k) = m_Balls(i).Vel(k) * -0.7
                m_Balls(i).Pos(k) = Sgn(m_Balls(i).Pos(k)) * ROOM_SIZE
            End If
        Next k

        ' Ball-ball collision
        For j = i + 1 To 4
            dx = m_Balls(i).Pos(0) - m_Balls(j).Pos(0)
            dy = m_Balls(i).Pos(1) - m_Balls(j).Pos(1)
            dz = m_Balls(i).Pos(2) - m_Balls(j).Pos(2)
            d = Sqr(dx * dx + dy * dy + dz * dz)

            If d < (BALL_RADIUS * 2) And d > 0.001 Then
                overlap = (BALL_RADIUS * 2) - d
                nx = dx / d: ny = dy / d: nz = dz / d

                ' Separate balls
                m_Balls(i).Pos(0) = m_Balls(i).Pos(0) + nx * (overlap / 2)
                m_Balls(i).Pos(1) = m_Balls(i).Pos(1) + ny * (overlap / 2)
                m_Balls(i).Pos(2) = m_Balls(i).Pos(2) + nz * (overlap / 2)
                m_Balls(j).Pos(0) = m_Balls(j).Pos(0) - nx * (overlap / 2)
                m_Balls(j).Pos(1) = m_Balls(j).Pos(1) - ny * (overlap / 2)
                m_Balls(j).Pos(2) = m_Balls(j).Pos(2) - nz * (overlap / 2)

                ' Relative velocity at point of contact
                relVx = m_Balls(i).Vel(0) - m_Balls(j).Vel(0)
                relVy = m_Balls(i).Vel(1) - m_Balls(j).Vel(1)
                relVz = m_Balls(i).Vel(2) - m_Balls(j).Vel(2)
                relSpeed = Sqr(relVx * relVx + relVy * relVy + relVz * relVz)

                ' Swap linear velocities (elastic collision)
                For k = 0 To 2
                    tv(k) = m_Balls(i).Vel(k)
                    m_Balls(i).Vel(k) = m_Balls(j).Vel(k)
                    m_Balls(j).Vel(k) = tv(k)
                Next k

                ' Transfer some relative speed into spin on both balls
                ' Spin axis is perpendicular to collision normal
                ' Cross product of normal with relative velocity gives spin axis
                Dim spinKick As Single
                spinKick = relSpeed * SPIN_TRANSFER

                ' Cross(normal, relVel) gives torque axis
                Dim txI As Single, tyI As Single, tzI As Single
                txI = ny * relVz - nz * relVy
                tyI = nz * relVx - nx * relVz
                tzI = nx * relVy - ny * relVx

                ' Apply spin kick to both balls (opposite directions)
                m_Balls(i).SpinX = m_Balls(i).SpinX + txI * spinKick
                m_Balls(i).SpinY = m_Balls(i).SpinY + tyI * spinKick
                m_Balls(i).SpinZ = m_Balls(i).SpinZ + tzI * spinKick

                m_Balls(j).SpinX = m_Balls(j).SpinX - txI * spinKick
                m_Balls(j).SpinY = m_Balls(j).SpinY - tyI * spinKick
                m_Balls(j).SpinZ = m_Balls(j).SpinZ - tzI * spinKick
            End If
        Next j
    Next i
End Sub

Public Sub DrawScene()
    Dim i As Long
    glClear GL_COLOR_BUFFER_BIT Or GL_DEPTH_BUFFER_BIT

    If m_ViewMode = 1 Then Draw360Sphere

    glMatrixMode GL_PROJECTION: glLoadIdentity
    glFrustum -1, 1, -0.75, 0.75, 1.5, 100
    glMatrixMode GL_MODELVIEW: glLoadIdentity
    glTranslatef 0, 0, ZoomZ
    glRotatef RotX, 1, 0, 0
    glRotatef RotY, 0, 1, 0

    ' Particles
    glDisable GL_LIGHTING
    glEnable GL_BLEND
    glBlendFunc GL_SRC_ALPHA, GL_ONE_MINUS_SRC_ALPHA
    glEnable GL_POINT_SMOOTH
    glPointSize 4!
    glBegin GL_POINTS
    For i = 0 To 99
        If m_Particles(i).Life > 0 Then
            glColor4f 0, 1, 1, m_Particles(i).Life
            glVertex3f m_Particles(i).Pos(0), m_Particles(i).Pos(1), m_Particles(i).Pos(2)
        End If
    Next i
    glEnd
    glDisable GL_POINT_SMOOTH
    glDisable GL_BLEND
    glEnable GL_LIGHTING

    ' Data bars
    glDisable GL_TEXTURE_2D
    glColor4f 0.5, 0.8, 1, 1
    For i = 1 To 10
        If m_DataBuffer(i) > 0.01 Then
            glPushMatrix
            glTranslatef (i - 5.5) * 1.5, m_DataBuffer(i) / 2, 5
            glScalef 0.5, m_DataBuffer(i), 0.5
            DrawCubeGeometry 1#
            glPopMatrix
        End If
    Next i

    ' Physics balls - now with per-ball spin rotation applied
    glEnable GL_TEXTURE_2D
    glColor4f 1, 1, 1, 1
    glBindTexture GL_TEXTURE_2D, m_CubeTexture
    For i = 0 To 4
        glPushMatrix
            glTranslatef m_Balls(i).Pos(0), m_Balls(i).Pos(1), m_Balls(i).Pos(2)
            ' Apply accumulated spin rotations so texture visibly rotates
            glRotatef m_Balls(i).AngX, 1, 0, 0
            glRotatef m_Balls(i).AngY, 0, 1, 0
            glRotatef m_Balls(i).AngZ, 0, 0, 1
            If m_ShapeMode = 0 Then
                DrawSphereGeometry BALL_RADIUS, 24, 24
            Else
                DrawCubeGeometry BALL_RADIUS
            End If
        glPopMatrix
    Next i

    Draw2DUI
End Sub

Private Sub Draw360Sphere()
    Dim i As Long, j As Long
    Dim lat0 As Double, z0 As Double, r0 As Double
    Dim lat1 As Double, z1 As Double, r1 As Double
    Dim lng As Double, x As Double, y As Double
    Dim rings As Long, sectors As Long
    rings = 32: sectors = 64

    glDisable GL_LIGHTING
    glDisable GL_DEPTH_TEST
    glEnable GL_TEXTURE_2D
    glBindTexture GL_TEXTURE_2D, m_SphereTexture
    glColor4f 1, 1, 1, 1

    glMatrixMode GL_PROJECTION: glLoadIdentity
    glFrustum -1, 1, -0.75, 0.75, 0.5, 200

    glMatrixMode GL_MODELVIEW: glLoadIdentity
    glRotatef m_SpherePitch, 1, 0, 0
    glRotatef m_SphereYaw, 0, 1, 0

    For i = 0 To rings - 1
        lat0 = PI * (-0.5 + CDbl(i) / rings)
        z0 = Sin(lat0): r0 = Cos(lat0)
        lat1 = PI * (-0.5 + CDbl(i + 1) / rings)
        z1 = Sin(lat1): r1 = Cos(lat1)

        glBegin GL_QUAD_STRIP
        For j = 0 To sectors
            lng = 2 * PI * CDbl(j) / sectors
            x = Cos(lng): y = Sin(lng)
            glTexCoord2f 1# - CDbl(j) / sectors, CDbl(i) / rings
            glVertex3f x * r0 * 50, z0 * 50, y * r0 * 50
            glTexCoord2f 1# - CDbl(j) / sectors, CDbl(i + 1) / rings
            glVertex3f x * r1 * 50, z1 * 50, y * r1 * 50
        Next j
        glEnd
    Next i

    glDisable GL_TEXTURE_2D
    glEnable GL_DEPTH_TEST
    glEnable GL_LIGHTING
End Sub

Private Sub DrawSphereGeometry(radius As Single, rings As Long, sectors As Long)
    Dim i As Long, j As Long
    Dim lat0 As Single, z0 As Single, r0 As Single
    Dim lat1 As Single, z1 As Single, r1 As Single
    Dim lng As Single, x As Single, y As Single

    For i = 0 To rings - 1
        lat0 = PI * (-0.5 + CSng(i) / rings)
        z0 = Sin(lat0): r0 = Cos(lat0)
        lat1 = PI * (-0.5 + CSng(i + 1) / rings)
        z1 = Sin(lat1): r1 = Cos(lat1)

        glBegin GL_QUAD_STRIP
        For j = 0 To sectors
            lng = 2 * PI * CSng(j) / sectors
            x = Cos(lng): y = Sin(lng)
            glNormal3f x * r1, z1, y * r1
            glTexCoord2f CSng(j) / sectors, CSng(i + 1) / rings
            glVertex3f x * r1 * radius, z1 * radius, y * r1 * radius
            glNormal3f x * r0, z0, y * r0
            glTexCoord2f CSng(j) / sectors, CSng(i) / rings
            glVertex3f x * r0 * radius, z0 * radius, y * r0 * radius
        Next j
        glEnd
    Next i
End Sub

Private Sub DrawCubeGeometry(size As Single)
    Dim s As Single: s = size
    glBegin GL_QUADS
        glNormal3f 0, 0, 1
        glTexCoord2f 0, 0: glVertex3f -s, -s, s
        glTexCoord2f 1, 0: glVertex3f s, -s, s
        glTexCoord2f 1, 1: glVertex3f s, s, s
        glTexCoord2f 0, 1: glVertex3f -s, s, s

        glNormal3f 0, 0, -1
        glTexCoord2f 1, 0: glVertex3f -s, -s, -s
        glTexCoord2f 1, 1: glVertex3f -s, s, -s
        glTexCoord2f 0, 1: glVertex3f s, s, -s
        glTexCoord2f 0, 0: glVertex3f s, -s, -s

        glNormal3f 0, 1, 0
        glTexCoord2f 0, 1: glVertex3f -s, s, -s
        glTexCoord2f 0, 0: glVertex3f -s, s, s
        glTexCoord2f 1, 0: glVertex3f s, s, s
        glTexCoord2f 1, 1: glVertex3f s, s, -s

        glNormal3f 0, -1, 0
        glTexCoord2f 1, 1: glVertex3f -s, -s, -s
        glTexCoord2f 0, 1: glVertex3f s, -s, -s
        glTexCoord2f 0, 0: glVertex3f s, -s, s
        glTexCoord2f 1, 0: glVertex3f -s, -s, s

        glNormal3f 1, 0, 0
        glTexCoord2f 1, 0: glVertex3f s, -s, -s
        glTexCoord2f 1, 1: glVertex3f s, s, -s
        glTexCoord2f 0, 1: glVertex3f s, s, s
        glTexCoord2f 0, 0: glVertex3f s, -s, s

        glNormal3f -1, 0, 0
        glTexCoord2f 0, 0: glVertex3f -s, -s, -s
        glTexCoord2f 1, 0: glVertex3f -s, -s, s
        glTexCoord2f 1, 1: glVertex3f -s, s, s
        glTexCoord2f 0, 1: glVertex3f -s, s, -s
    glEnd
End Sub

Private Sub Draw2DUI()
    glDisable GL_LIGHTING
    glDisable GL_DEPTH_TEST
    glEnable GL_TEXTURE_2D
    glEnable GL_BLEND
    glBlendFunc GL_SRC_ALPHA, GL_ONE_MINUS_SRC_ALPHA
    glColor4f 1, 1, 1, 1

    glMatrixMode GL_PROJECTION: glPushMatrix: glLoadIdentity
    glOrtho 0, 1024, 768, 0, -1, 1
    glMatrixMode GL_MODELVIEW: glPushMatrix: glLoadIdentity

    glBindTexture GL_TEXTURE_2D, fpsTexture
    DrawUIButton 10, 10, 128, 32

    glBindTexture GL_TEXTURE_2D, m_ResetTexture
    DrawUIButton BTN_X, GetButtonY(0), BTN_W, BTN_HEIGHT

    glBindTexture GL_TEXTURE_2D, m_Btn360
    DrawUIButton BTN_X, GetButtonY(1), BTN_W, BTN_HEIGHT

    glBindTexture GL_TEXTURE_2D, m_BtnShape
    DrawUIButton BTN_X, GetButtonY(2), BTN_W, BTN_HEIGHT

    glMatrixMode GL_PROJECTION: glPopMatrix
    glMatrixMode GL_MODELVIEW: glPopMatrix

    glDisable GL_BLEND
    glDisable GL_TEXTURE_2D
    glEnable GL_DEPTH_TEST
    glEnable GL_LIGHTING
End Sub

Private Sub DrawUIButton(x As Long, y As Long, w As Long, h As Long)
    glBegin GL_QUADS
        glTexCoord2f 0, 0: glVertex2f x, y
        glTexCoord2f 1, 0: glVertex2f x + w, y
        glTexCoord2f 1, 1: glVertex2f x + w, y + h
        glTexCoord2f 0, 1: glVertex2f x, y + h
    glEnd
End Sub

Public Function HandleMessage(ByVal h As LongPtr, ByVal m As Long, ByVal w As LongPtr, ByVal l As LongPtr) As LongPtr
    On Error Resume Next
    Dim mx As Long, my As Long
    mx = CLng(l And &HFFFF&)
    my = CLng((l \ &H10000) And &HFFFF&)

    Select Case m
        Case WM_MOUSEMOVE
            If mx >= BTN_X And mx <= BTN_X + BTN_W And _
               my >= UI_START_Y And my <= GetButtonY(2) + BTN_HEIGHT Then
                SetCursor LoadCursor(0, IDC_HAND)
            End If
            If m_IsDragging Then
                If m_ViewMode = 0 Then
                    RotY = RotY + (mx - m_LastX) * 0.2
                    RotX = RotX + (my - m_LastY) * 0.2
                Else
                    m_SphereYaw = m_SphereYaw + (mx - m_LastX) * 0.3
                    m_SpherePitch = m_SpherePitch + (my - m_LastY) * 0.3
                    If m_SpherePitch > 85 Then m_SpherePitch = 85
                    If m_SpherePitch < -85 Then m_SpherePitch = -85
                End If
                m_LastX = mx: m_LastY = my
            End If

        Case WM_LBUTTONDOWN
            If mx >= BTN_X And mx <= BTN_X + BTN_W Then
                If my >= GetButtonY(0) And my <= GetButtonY(0) + BTN_HEIGHT Then
                    InitBalls
                ElseIf my >= GetButtonY(1) And my <= GetButtonY(1) + BTN_HEIGHT Then
                    m_ViewMode = IIf(m_ViewMode = 0, 1, 0)
                ElseIf my >= GetButtonY(2) And my <= GetButtonY(2) + BTN_HEIGHT Then
                    m_ShapeMode = IIf(m_ShapeMode = 0, 1, 0)
                End If
            Else
                m_IsDragging = True
                m_LastX = mx: m_LastY = my
            End If

        Case WM_LBUTTONUP
            m_IsDragging = False

        Case WM_MOUSEWHEEL
            Dim delta As Long
            delta = CLng(w \ &H10000)
            If m_ViewMode = 0 Then
                ZoomZ = ZoomZ + (delta / 1200#)
            Else
                m_SphereFOV = m_SphereFOV - (delta / 120#) * 5#
                If m_SphereFOV < 20 Then m_SphereFOV = 20
                If m_SphereFOV > 120 Then m_SphereFOV = 120
            End If

        Case WM_CLOSE
            IsRunning = False
            DestroyWindow h
            HandleMessage = 0
            Exit Function

        Case WM_DESTROY
            PostQuitMessage 0
            HandleMessage = 0
            Exit Function

        Case Else
            HandleMessage = DefWindowProc(h, m, w, l)
    End Select
End Function

Public Function Load360Texture(ByVal path As String) As Boolean
    Dim img As LongPtr, imgW As Long, imgH As Long
    Dim bD As BitmapData, texID As Long
    If GdipLoadImageFromFile(StrPtr(path), img) = 0 Then
        GdipGetImageWidth img, imgW
        GdipGetImageHeight img, imgH
        GdipBitmapLockBits img, ByVal 0&, ImageLockModeRead, PixelFormat32bppARGB, bD
        glGenTextures 1, texID
        glBindTexture GL_TEXTURE_2D, texID
        glTexImage2D GL_TEXTURE_2D, 0, GL_RGBA, imgW, imgH, 0, GL_BGRA, GL_UNSIGNED_BYTE, ByVal bD.Scan0
        glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_MIN_FILTER, GL_LINEAR
        glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_MAG_FILTER, GL_LINEAR
        GdipBitmapUnlockBits img, bD
        GdipDisposeImage img
        m_SphereTexture = texID
        Load360Texture = True
    Else
        MsgBox "Failed to load 360 image: " & path
    End If
End Function

Public Function LoadTextureFromFile(ByVal filePath As String) As Long
    Dim fullPath As String
    If Mid(filePath, 2, 1) = ":" Then
        fullPath = filePath
    Else
        fullPath = ThisWorkbook.path & "\" & filePath
    End If

    Dim img As LongPtr, imgW As Long, imgH As Long
    Dim bD As BitmapData, texID As Long

    If GdipLoadImageFromFile(StrPtr(fullPath), img) = 0 Then
        GdipGetImageWidth img, imgW
        GdipGetImageHeight img, imgH
        GdipBitmapLockBits img, ByVal 0&, ImageLockModeRead, PixelFormat32bppARGB, bD
        glGenTextures 1, texID
        glBindTexture GL_TEXTURE_2D, texID
        glTexImage2D GL_TEXTURE_2D, 0, GL_RGBA, imgW, imgH, 0, GL_BGRA, GL_UNSIGNED_BYTE, ByVal bD.Scan0
        glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_MIN_FILTER, GL_LINEAR
        glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_MAG_FILTER, GL_LINEAR
        GdipBitmapUnlockBits img, bD
        GdipDisposeImage img
        LoadTextureFromFile = texID
    Else
        MsgBox "Failed to load image: " & fullPath
        LoadTextureFromFile = 0
    End If
End Function

Public Sub UpdateButtonTexture(texID As Long, txt As String, clr As Long)
    Dim g As LongPtr, r As RECTF, bD As BitmapData, bmp As LongPtr
    GdipCreateBitmapFromScan0 128, 64, 0, PixelFormat32bppARGB, ByVal 0&, bmp
    GdipGetImageGraphicsContext bmp, g
    GdipGraphicsClear g, clr
    r.Left = 5: r.Top = 15: r.Right = 120: r.Bottom = 50
    GdipDrawString g, StrPtr(txt), -1, font, r, 0, brush
    GdipBitmapLockBits bmp, ByVal 0&, ImageLockModeRead, PixelFormat32bppARGB, bD
    glBindTexture GL_TEXTURE_2D, texID
    glTexImage2D GL_TEXTURE_2D, 0, GL_RGBA, 128, 64, 0, GL_BGRA, GL_UNSIGNED_BYTE, ByVal bD.Scan0
    glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_MIN_FILTER, GL_LINEAR
    glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_MAG_FILTER, GL_LINEAR
    GdipBitmapUnlockBits bmp, bD
    GdipDeleteGraphics g
    GdipDisposeImage bmp
End Sub

Public Sub ProcessMessages()
    Dim msg As msg
    If PeekMessage(msg, 0, 0, 0, PM_REMOVE) Then
        TranslateMessage msg
        DispatchMessage msg
    End If
End Sub

Public Sub DrainMessages()
    Dim msg As msg
    Dim i As Long
    For i = 1 To 100
        If PeekMessage(msg, 0, 0, 0, PM_REMOVE) Then
            TranslateMessage msg
            DispatchMessage msg
        Else
            Exit For
        End If
    Next i
End Sub

Public Sub DoSwapBuffers()
    SwapBuffers hDC
End Sub

Public Sub CleanupDemo()
    On Error Resume Next
    If font <> 0 Then GdipDeleteFont font
    If brush <> 0 Then GdipDeleteBrush brush
    If token <> 0 Then GdiplusShutdown token
    If hRC <> 0 Then
        glDeleteTextures 1, fpsTexture
        glDeleteTextures 1, m_ResetTexture
        glDeleteTextures 1, m_Btn360
        glDeleteTextures 1, m_BtnShape
        If m_SphereTexture <> 0 Then glDeleteTextures 1, m_SphereTexture
        If m_CubeTexture <> 0 Then glDeleteTextures 1, m_CubeTexture
        If m_PetTexture <> 0 Then glDeleteTextures 1, m_PetTexture
    End If
End Sub

Public Sub Destroy()
    On Error Resume Next
    IsRunning = False
    Set g_OpenGLWindow = Nothing
    If hRC <> 0 Then
        wglMakeCurrent 0, 0
        wglDeleteContext hRC
        hRC = 0
    End If
    If hDC <> 0 Then
        ReleaseDC hWnd, hDC
        hDC = 0
    End If
    If hWnd <> 0 Then
        DestroyWindow hWnd
        hWnd = 0
    End If
    UnregisterClass "VBA_PhysX_Class", GetModuleHandle(vbNullString)
End Sub

