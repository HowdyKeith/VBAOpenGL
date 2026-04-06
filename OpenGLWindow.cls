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

' --- UI Constants ---
Private Const BTN_X As Long = 10
Private Const BTN_W As Long = 128
Private Const BTN_HEIGHT As Long = 35
Private Const BTN_GAP As Long = 10 ' <--- ADJUST THIS FOR SPACING
Private Const UI_START_Y As Long = 50

' --- Physics Constants ---
Private Const BALL_RADIUS As Single = 1.2
Private Const ROOM_SIZE As Single = 8#
Private Const FRICTION As Single = 0.98

' --- Interaction & Core ---
Private m_LastX As Long, m_LastY As Long
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
Private m_SphereYaw As Double, m_SpherePitch As Double, m_SphereFOV As Double
Private token As LongPtr, font As LongPtr, brush As LongPtr
Private fpsTexture As Long, m_ResetTexture As Long, m_Btn360 As Long, m_BtnShape As Long
Private m_PetFilePath As String

' --- Physics Entities ---
Private Type PhysicalObject: Pos(2) As Single: Vel(2) As Single: End Type
Private m_Balls(0 To 4) As PhysicalObject

' --- Helper for UI Hitbox Math ---
Private Function GetButtonY(index As Long) As Long
    GetButtonY = UI_START_Y + (index * (BTN_HEIGHT + BTN_GAP))
End Function

Public Sub SetPetPath(ByVal path As String): m_PetFilePath = path: End Sub

Public Function Create(Width As Long, Height As Long, Title As String) As Boolean
    Dim wc As WNDCLASSEX, hInst As LongPtr: hInst = GetModuleHandle(vbNullString)
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
    hWnd = CreateWindowEx(0, wc.lpszClassName, Title, WS_OVERLAPPEDWINDOW Or WS_CLIPCHILDREN Or WS_CLIPSIBLINGS, CW_USEDEFAULT, CW_USEDEFAULT, Width, Height, 0, 0, hInst, 0)
    If hWnd <> 0 Then
        hDC = GetDC(hWnd): SetupPixelFormat: hRC = wglCreateContext(hDC): wglMakeCurrent hDC, hRC
        ShowWindow hWnd, SW_SHOW: IsRunning = True: ZoomZ = -20#: Create = True
        m_ViewMode = 1
    End If
End Function

Private Sub SetupPixelFormat()
    Dim pfd As PIXELFORMATDESCRIPTOR: pfd.nSize = LenB(pfd): pfd.nVersion = 1
    pfd.dwFlags = PFD_DRAW_TO_WINDOW Or PFD_SUPPORT_OPENGL Or PFD_DOUBLEBUFFER: pfd.iPixelType = PFD_TYPE_RGBA: pfd.cColorBits = 32: pfd.cDepthBits = 24
    SetPixelFormat hDC, ChoosePixelFormat(hDC, pfd), pfd
End Sub

Public Sub InitDemo()
    Dim gpInput As GDIPlusStartupInput: gpInput.GdiplusVersion = 1: GdiplusStartup token, gpInput, ByVal 0&
    Dim family As LongPtr: GdipCreateFontFamilyFromName StrPtr("Consolas"), 0, family: GdipCreateFont family, 11, 0, 3, font: GdipCreateSolidFill &HFFFFFFFF, brush: GdipDeleteFontFamily family
    glGenTextures 1, fpsTexture: glGenTextures 1, m_ResetTexture: glGenTextures 1, m_Btn360: glGenTextures 1, m_BtnShape
    glEnable GL_LIGHTING: glEnable GL_LIGHT0: glEnable GL_DEPTH_TEST: glShadeModel GL_SMOOTH
    
    Dim i As Long: For i = 0 To 4: m_Balls(i).Pos(0) = (i - 2) * 2: m_Balls(i).Vel(0) = (Rnd - 0.5) * 0.2: Next
    m_SphereFOV = 80#
    UpdateButtonTexture m_ResetTexture, "RESET PHYSICS", &HFF444444
    UpdateButtonTexture m_Btn360, "TOGGLE SKYBOX", &HFF664422
    UpdateButtonTexture m_BtnShape, "TOGGLE SHAPE", &HFF224466
End Sub

Public Sub UpdatePhysics()
    Dim i As Long, j As Long, k As Long
    Dim d As Single, dx As Single, dy As Single, dz As Single
    Dim overlap As Single, nx As Single, ny As Single, nz As Single
    Dim tv(2) As Single
    
    For i = 0 To 4
        For k = 0 To 2: m_Balls(i).Vel(k) = m_Balls(i).Vel(k) * FRICTION: Next
        For k = 0 To 2: m_Balls(i).Pos(k) = m_Balls(i).Pos(k) + m_Balls(i).Vel(k): Next
        
        For k = 0 To 2
            If Abs(m_Balls(i).Pos(k)) > ROOM_SIZE Then
                m_Balls(i).Vel(k) = m_Balls(i).Vel(k) * -0.7
                m_Balls(i).Pos(k) = Sgn(m_Balls(i).Pos(k)) * ROOM_SIZE
            End If
        Next
        
        For j = i + 1 To 4
            dx = m_Balls(i).Pos(0) - m_Balls(j).Pos(0): dy = m_Balls(i).Pos(1) - m_Balls(j).Pos(1): dz = m_Balls(i).Pos(2) - m_Balls(j).Pos(2)
            d = Sqr(dx * dx + dy * dy + dz * dz)
            If d < (BALL_RADIUS * 2) And d > 0.001 Then
                overlap = (BALL_RADIUS * 2) - d
                nx = dx / d: ny = dy / d: nz = dz / d
                m_Balls(i).Pos(0) = m_Balls(i).Pos(0) + nx * (overlap / 2): m_Balls(i).Pos(1) = m_Balls(i).Pos(1) + ny * (overlap / 2): m_Balls(i).Pos(2) = m_Balls(i).Pos(2) + nz * (overlap / 2)
                m_Balls(j).Pos(0) = m_Balls(j).Pos(0) - nx * (overlap / 2): m_Balls(j).Pos(1) = m_Balls(j).Pos(1) - ny * (overlap / 2): m_Balls(j).Pos(2) = m_Balls(j).Pos(2) - nz * (overlap / 2)
                For k = 0 To 2: tv(k) = m_Balls(i).Vel(k): m_Balls(i).Vel(k) = m_Balls(j).Vel(k): m_Balls(j).Vel(k) = tv(k): Next
            End If
        Next
    Next
End Sub

Public Sub DrawScene()
    glClear GL_COLOR_BUFFER_BIT Or GL_DEPTH_BUFFER_BIT
    If m_ViewMode = 1 Then Draw360Sphere
    
    glMatrixMode GL_PROJECTION: glLoadIdentity: glFrustum -1, 1, -0.75, 0.75, 1.5, 100
    glMatrixMode GL_MODELVIEW: glLoadIdentity: glTranslatef 0, 0, ZoomZ: glRotatef RotX, 1, 0, 0: glRotatef RotY, 0, 1, 0
    
    glEnable GL_TEXTURE_2D: glColor4f 1, 1, 1, 1
    glBindTexture GL_TEXTURE_2D, m_CubeTexture
    
    Dim i As Long: For i = 0 To 4
glPushMatrix:         glTranslatef m_Balls(i).Pos(0), m_Balls(i).Pos(1), m_Balls(i).Pos(2)
        If m_ShapeMode = 0 Then
            DrawSphereGeometry BALL_RADIUS, 24, 24
        Else
            DrawCubeGeometry BALL_RADIUS
        End If
        glPopMatrix
    Next
    Draw2DUI
End Sub

Private Sub Draw360Sphere()
    Const PI As Double = 3.14159265358979: Const stacks As Long = 32: Const slices As Long = 64: Const RADIUS As Double = 30#
    glMatrixMode GL_PROJECTION: glLoadIdentity: Dim t As Double: t = 0.1 * Tan((m_SphereFOV * PI / 180#) / 2): glFrustum -t * (4 / 3), t * (4 / 3), -t, t, 0.1, 100
    glMatrixMode GL_MODELVIEW: glLoadIdentity: glRotatef m_SpherePitch, 1, 0, 0: glRotatef m_SphereYaw, 0, 1, 0
    glDisable GL_LIGHTING: glDisable GL_DEPTH_TEST: glEnable GL_CULL_FACE: glCullFace GL_FRONT: glEnable GL_TEXTURE_2D
    glBindTexture GL_TEXTURE_2D, m_SphereTexture: glColor4f 1, 1, 1, 1
    Dim i As Long, j As Long, lat0 As Double, lat1 As Double, lng As Double: For i = 0 To stacks - 1: lat0 = PI * (-0.5 + i / stacks): lat1 = PI * (-0.5 + (i + 1) / stacks): glBegin GL_TRIANGLE_STRIP: For j = 0 To slices: lng = 2 * PI * j / slices: glTexCoord2d 1 - (j / slices), 1 - (i / stacks): glVertex3d RADIUS * Cos(lat0) * Cos(lng), RADIUS * Sin(lat0), RADIUS * Cos(lat0) * Sin(lng): glTexCoord2d 1 - (j / slices), 1 - ((i + 1) / stacks): glVertex3d RADIUS * Cos(lat1) * Cos(lng), RADIUS * Sin(lat1), RADIUS * Cos(lat1) * Sin(lng): Next: glEnd: Next
    glDisable GL_CULL_FACE: glEnable GL_DEPTH_TEST: glEnable GL_LIGHTING
End Sub

Private Sub DrawSphereGeometry(r As Single, slices As Long, stacks As Long)
    Dim i As Long, j As Long, PI As Double: PI = 3.14159265358979: glDisable GL_TEXTURE_GEN_S: glDisable GL_TEXTURE_GEN_T
    For i = 0 To stacks - 1
        Dim lat0 As Double: lat0 = PI * (-0.5 + i / stacks): Dim lat1 As Double: lat1 = PI * (-0.5 + (i + 1) / stacks)
        glBegin GL_TRIANGLE_STRIP
        For j = 0 To slices
            Dim lng As Double: lng = 2 * PI * j / slices: Dim x As Double: x = Cos(lng): Dim y As Double: y = Sin(lng)
            glTexCoord2f CSng(j / slices), CSng(i / stacks): glNormal3f CSng(x * Cos(lat0)), CSng(Sin(lat0)), CSng(y * Cos(lat0)): glVertex3f CSng(r * x * Cos(lat0)), CSng(r * Sin(lat0)), CSng(r * y * Cos(lat0))
            glTexCoord2f CSng(j / slices), CSng((i + 1) / stacks): glNormal3f CSng(x * Cos(lat1)), CSng(Sin(lat1)), CSng(y * Cos(lat1)): glVertex3f CSng(r * x * Cos(lat1)), CSng(r * Sin(lat1)), CSng(r * y * Cos(lat1))
        Next
        glEnd
    Next
End Sub

Private Sub DrawCubeGeometry(s As Single)
    glBegin GL_QUADS
    glNormal3f 0, 0, 1: glTexCoord2f 0, 0: glVertex3f -s, -s, s: glTexCoord2f 1, 0: glVertex3f s, -s, s: glTexCoord2f 1, 1: glVertex3f s, s, s: glTexCoord2f 0, 1: glVertex3f -s, s, s
    glNormal3f 0, 0, -1: glTexCoord2f 1, 0: glVertex3f -s, -s, -s: glTexCoord2f 1, 1: glVertex3f -s, s, -s: glTexCoord2f 0, 1: glVertex3f s, s, -s: glTexCoord2f 0, 0: glVertex3f s, -s, -s
    glNormal3f 0, 1, 0: glTexCoord2f 0, 1: glVertex3f -s, s, -s: glTexCoord2f 0, 0: glVertex3f -s, s, s: glTexCoord2f 1, 0: glVertex3f s, s, s: glTexCoord2f 1, 1: glVertex3f s, s, -s
    glNormal3f 0, -1, 0: glTexCoord2f 1, 1: glVertex3f -s, -s, -s: glTexCoord2f 0, 1: glVertex3f s, -s, -s: glTexCoord2f 0, 0: glVertex3f s, -s, s: glTexCoord2f 1, 0: glVertex3f -s, -s, s
    glNormal3f 1, 0, 0: glTexCoord2f 1, 0: glVertex3f s, -s, -s: glTexCoord2f 1, 1: glVertex3f s, s, -s: glTexCoord2f 0, 1: glVertex3f s, s, s: glTexCoord2f 0, 0: glVertex3f s, -s, s
    glNormal3f -1, 0, 0: glTexCoord2f 0, 0: glVertex3f -s, -s, -s: glTexCoord2f 1, 0: glVertex3f -s, -s, s: glTexCoord2f 1, 1: glVertex3f -s, s, s: glTexCoord2f 0, 1: glVertex3f -s, s, -s
    glEnd
End Sub

Private Sub Draw2DUI()
    glDisable GL_LIGHTING: glDisable GL_DEPTH_TEST: glEnable GL_BLEND: glBlendFunc GL_SRC_ALPHA, GL_ONE_MINUS_SRC_ALPHA
    glMatrixMode GL_PROJECTION: glLoadIdentity: glOrtho 0, 800, 600, 0, -1, 1: glMatrixMode GL_MODELVIEW: glLoadIdentity
    DrawUIButton BTN_X, GetButtonY(0), m_ResetTexture
    DrawUIButton BTN_X, GetButtonY(1), m_Btn360
    DrawUIButton BTN_X, GetButtonY(2), m_BtnShape
    glBindTexture GL_TEXTURE_2D, fpsTexture: glBegin GL_QUADS: glTexCoord2f 0, 0: glVertex2f 10, 10: glTexCoord2f 1, 0: glVertex2f 138, 10: glTexCoord2f 1, 1: glVertex2f 138, 42: glTexCoord2f 0, 1: glVertex2f 10, 42: glEnd
End Sub

Private Sub DrawUIButton(x As Single, y As Single, texID As Long)
    glBindTexture GL_TEXTURE_2D, texID: glBegin GL_QUADS: glTexCoord2f 0, 0: glVertex2f x, y: glTexCoord2f 1, 0: glVertex2f x + BTN_W, y: glTexCoord2f 1, 1: glVertex2f x + BTN_W, y + BTN_HEIGHT: glTexCoord2f 0, 1: glVertex2f x, y + BTN_HEIGHT: glEnd
End Sub

Public Function HandleMessage(ByVal h As LongPtr, ByVal m As Long, ByVal w As LongPtr, ByVal l As LongPtr) As LongPtr
    On Error Resume Next: Dim mx As Long, my As Long: mx = CLng(l And &HFFFF&): my = CLng((l \ &H10000) And &HFFFF&)
    Select Case m
        Case WM_MOUSEMOVE
            ' Cursor Hover logic based on constants
            If mx >= BTN_X And mx <= BTN_X + BTN_W And my >= UI_START_Y And my <= GetButtonY(2) + BTN_HEIGHT Then
                SetCursor LoadCursor(0, IDC_HAND)
            End If
            If m_IsDragging Then
                If m_ViewMode = 0 Then
                    RotY = RotY + (mx - m_LastX) * 0.4: RotX = RotX + (my - m_LastY) * 0.4
                Else
                    m_SphereYaw = m_SphereYaw + (mx - m_LastX) * 0.1: m_SpherePitch = m_SpherePitch + (my - m_LastY) * 0.1
                End If
                m_LastX = mx: m_LastY = my
            End If
        Case WM_LBUTTONDOWN
            If mx >= BTN_X And mx <= BTN_X + BTN_W Then
                If my >= GetButtonY(0) And my <= GetButtonY(0) + BTN_HEIGHT Then
                    Dim i As Long: For i = 0 To 4: m_Balls(i).Pos(0) = (i - 2) * 2: m_Balls(i).Vel(0) = (Rnd - 0.5) * 0.5: Next: m_IsDragging = False
                ElseIf my >= GetButtonY(1) And my <= GetButtonY(1) + BTN_HEIGHT Then
                    m_ViewMode = IIf(m_ViewMode = 0, 1, 0): m_IsDragging = False
                ElseIf my >= GetButtonY(2) And my <= GetButtonY(2) + BTN_HEIGHT Then
                    m_ShapeMode = IIf(m_ShapeMode = 0, 1, 0): m_IsDragging = False
                End If
            Else
                m_IsDragging = True: m_LastX = mx: m_LastY = my
            End If
            ' Hotspot Logic
            If m_ViewMode = 1 And m_IsDragging Then
                If (m_SphereYaw > 170 And m_SphereYaw < 190) Then
                    If m_PetTexture = 0 Then LoadPetTexture m_PetFilePath
                    m_CubeTexture = m_PetTexture: m_IsDragging = False
                End If
            End If
        Case WM_LBUTTONUP: m_IsDragging = False
        Case WM_MOUSEWHEEL: If m_ViewMode = 0 Then ZoomZ = ZoomZ + (CLng(w \ &H10000) / 1200) Else m_SphereFOV = m_SphereFOV - (CLng(w \ &H10000) / 1200) * 5
        Case WM_CLOSE: IsRunning = False: Case WM_DESTROY: PostQuitMessage 0: Case Else: HandleMessage = DefWindowProc(h, m, w, l)
    End Select
End Function

Public Function Load360Texture(ByVal path As String) As Boolean
    Dim img As LongPtr, w As Long, h As Long, bD As BitmapData, texID As Long
    If GdipLoadImageFromFile(StrPtr(path), img) = 0 Then
        GdipGetImageWidth img, w: GdipGetImageHeight img, h: GdipBitmapLockBits img, ByVal 0&, ImageLockModeRead, PixelFormat32bppARGB, bD
        glGenTextures 1, texID: glBindTexture GL_TEXTURE_2D, texID: glTexImage2D GL_TEXTURE_2D, 0, GL_RGBA, w, h, 0, GL_BGRA, GL_UNSIGNED_BYTE, ByVal bD.Scan0: glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_MIN_FILTER, GL_LINEAR: glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_MAG_FILTER, GL_LINEAR
        GdipBitmapUnlockBits img, bD: GdipDisposeImage img: m_SphereTexture = texID: Load360Texture = True
    End If
End Function

Public Function LoadPetTexture(ByVal path As String) As Boolean
    Dim img As LongPtr, w As Long, h As Long, bD As BitmapData, texID As Long
    If GdipLoadImageFromFile(StrPtr(path), img) = 0 Then
        GdipGetImageWidth img, w: GdipGetImageHeight img, h: GdipBitmapLockBits img, ByVal 0&, ImageLockModeRead, PixelFormat32bppARGB, bD
        glGenTextures 1, texID: glBindTexture GL_TEXTURE_2D, texID: glTexImage2D GL_TEXTURE_2D, 0, GL_RGBA, w, h, 0, GL_BGRA, GL_UNSIGNED_BYTE, ByVal bD.Scan0: glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_MIN_FILTER, GL_LINEAR: glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_MAG_FILTER, GL_LINEAR
        GdipBitmapUnlockBits img, bD: GdipDisposeImage img: m_PetTexture = texID: LoadPetTexture = True
    End If
End Function

Public Sub SetCubeTexture(texID As Long): m_CubeTexture = texID: End Sub
Public Function LoadTextureFromFile(ByVal fn As String) As Long
    ' Check if the texture loads successfully
    If LoadPetTexture(fn) Then
        ' Assign the class-level texture ID to the function result
        LoadTextureFromFile = m_PetTexture
    End If
End Function
Public Sub UpdateButtonTexture(texID As Long, txt As String, clr As Long)
    Dim g As LongPtr, r As RECTF, bD As BitmapData, bmp As LongPtr
    GdipCreateBitmapFromScan0 128, 64, 0, PixelFormat32bppARGB, ByVal 0&, bmp: GdipGetImageGraphicsContext bmp, g: GdipGraphicsClear g, clr
    r.Left = 5: r.Top = 15: r.Right = 120: r.Bottom = 50
    GdipDrawString g, StrPtr(txt), -1, font, r, 0, brush
    GdipBitmapLockBits bmp, ByVal 0&, ImageLockModeRead, PixelFormat32bppARGB, bD
    glBindTexture GL_TEXTURE_2D, texID: glTexImage2D GL_TEXTURE_2D, 0, GL_RGBA, 128, 64, 0, GL_BGRA, GL_UNSIGNED_BYTE, ByVal bD.Scan0: glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_MIN_FILTER, GL_LINEAR
    GdipBitmapUnlockBits bmp, bD: GdipDeleteGraphics g: GdipDisposeImage bmp
End Sub

Public Sub UpdateFPSOverlay(txt As String): UpdateButtonTexture fpsTexture, txt, &HFF000000: End Sub
Public Sub DoSwapBuffers(): SwapBuffers hDC: End Sub
Public Sub ProcessMessages(): Dim m As msg: Do While PeekMessage(m, 0, 0, 0, PM_REMOVE) <> 0: TranslateMessage m: DispatchMessage m: Loop: End Sub
Public Sub DrainMessages(): Dim m As msg: Do While PeekMessage(m, 0, 0, 0, PM_REMOVE) <> 0: Loop: End Sub

Public Sub CleanupDemo()
    If font <> 0 Then GdipDeleteFont font
    If brush <> 0 Then GdipDeleteBrush brush
    If token <> 0 Then GdiplusShutdown token
    If hRC <> 0 Then
        glDeleteTextures 1, fpsTexture: glDeleteTextures 1, m_ResetTexture: glDeleteTextures 1, m_Btn360: glDeleteTextures 1, m_BtnShape
        If m_PetTexture <> 0 Then glDeleteTextures 1, m_PetTexture
        If m_SphereTexture <> 0 Then glDeleteTextures 1, m_SphereTexture
    End If
End Sub

Public Sub Destroy()
    IsRunning = False: Set g_OpenGLWindow = Nothing
    If hRC <> 0 Then wglMakeCurrent 0, 0: wglDeleteContext hRC: If hDC <> 0 Then ReleaseDC hWnd, hDC: If hWnd <> 0 Then DestroyWindow hWnd
    UnregisterClass "VBA_PhysX_Class", GetModuleHandle(vbNullString)
End Sub
