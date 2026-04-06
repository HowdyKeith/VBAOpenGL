Attribute VB_Name = "ModWin32"
Option Explicit

' =============================================
' ModWin32.bas - Engine v1.45
' =============================================
Public g_OpenGLWindow As OpenGLWindow

#If VBA7 Then
    ' Windowing & Input
    Public Declare PtrSafe Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As LongPtr
    Public Declare PtrSafe Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As LongPtr, ByVal lpCursorName As Long) As LongPtr
    Public Declare PtrSafe Function SetCursor Lib "user32" (ByVal hCursor As LongPtr) As LongPtr
    Public Declare PtrSafe Function RegisterClassEx Lib "user32" Alias "RegisterClassExA" (pcWndClassEx As WNDCLASSEX) As Integer
    Public Declare PtrSafe Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As LongPtr, ByVal hMenu As LongPtr, ByVal hInstance As LongPtr, ByVal lpParam As LongPtr) As LongPtr
    Public Declare PtrSafe Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hWnd As LongPtr, ByVal uMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
    Public Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
    Public Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hWnd As LongPtr, ByVal hDC As LongPtr) As Long
    Public Declare PtrSafe Function ChoosePixelFormat Lib "gdi32" (ByVal hDC As LongPtr, pfd As PIXELFORMATDESCRIPTOR) As Long
    Public Declare PtrSafe Function SetPixelFormat Lib "gdi32" (ByVal hDC As LongPtr, ByVal iPixelFormat As Long, pfd As PIXELFORMATDESCRIPTOR) As Long
    Public Declare PtrSafe Function SwapBuffers Lib "gdi32" (ByVal hDC As LongPtr) As Long
    Public Declare PtrSafe Function DestroyWindow Lib "user32" (ByVal hWnd As LongPtr) As Long
    Public Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal nCmdShow As Long) As Long
    Public Declare PtrSafe Function PeekMessage Lib "user32" Alias "PeekMessageA" (lpMsg As msg, ByVal hWnd As LongPtr, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
    Public Declare PtrSafe Function TranslateMessage Lib "user32" (lpMsg As msg) As Long
    Public Declare PtrSafe Function DispatchMessage Lib "user32" Alias "DispatchMessageA" (lpMsg As msg) As LongPtr
    Public Declare PtrSafe Sub PostQuitMessage Lib "user32" (ByVal nExitCode As Long)
    Public Declare PtrSafe Function UnregisterClass Lib "user32" Alias "UnregisterClassA" (ByVal lpClassName As String, ByVal hInstance As LongPtr) As Long
    ' OpenGL Core
    Public Declare PtrSafe Function wglCreateContext Lib "opengl32" (ByVal hDC As LongPtr) As LongPtr
    Public Declare PtrSafe Function wglMakeCurrent Lib "opengl32" (ByVal hDC As LongPtr, ByVal hRC As LongPtr) As Long
    Public Declare PtrSafe Function wglDeleteContext Lib "opengl32" (ByVal hRC As LongPtr) As Long
    Public Declare PtrSafe Sub glEnable Lib "opengl32" (ByVal cap As Long)
    Public Declare PtrSafe Sub glDisable Lib "opengl32" (ByVal cap As Long)
    Public Declare PtrSafe Sub glClear Lib "opengl32" (ByVal mask As Long)
    Public Declare PtrSafe Sub glMatrixMode Lib "opengl32" (ByVal mode As Long)
    Public Declare PtrSafe Sub glLoadIdentity Lib "opengl32" ()
    Public Declare PtrSafe Sub glFrustum Lib "opengl32" (ByVal Left As Double, ByVal Right As Double, ByVal Bottom As Double, ByVal Top As Double, ByVal zNear As Double, ByVal zFar As Double)
    Public Declare PtrSafe Sub glOrtho Lib "opengl32" (ByVal Left As Double, ByVal Right As Double, ByVal Bottom As Double, ByVal Top As Double, ByVal zNear As Double, ByVal zFar As Double)
    Public Declare PtrSafe Sub glTranslatef Lib "opengl32" (ByVal x As Single, ByVal y As Single, ByVal z As Single)
    Public Declare PtrSafe Sub glRotatef Lib "opengl32" (ByVal Angle As Single, ByVal x As Single, ByVal y As Single, ByVal z As Single)
    Public Declare PtrSafe Sub glBegin Lib "opengl32" (ByVal mode As Long)
    Public Declare PtrSafe Sub glEnd Lib "opengl32" ()
    Public Declare PtrSafe Sub glColor4f Lib "opengl32" (ByVal red As Single, ByVal green As Single, ByVal blue As Single, ByVal alpha As Single)
    Public Declare PtrSafe Sub glVertex3f Lib "opengl32" (ByVal x As Single, ByVal y As Single, ByVal z As Single)
    Public Declare PtrSafe Sub glVertex3d Lib "opengl32" (ByVal x As Double, ByVal y As Double, ByVal z As Double)
    Public Declare PtrSafe Sub glVertex2f Lib "opengl32" (ByVal x As Single, ByVal y As Single)
    Public Declare PtrSafe Sub glNormal3f Lib "opengl32" (ByVal nx As Single, ByVal ny As Single, ByVal nz As Single)
    Public Declare PtrSafe Sub glTexCoord2f Lib "opengl32" (ByVal s As Single, ByVal t As Single)
    Public Declare PtrSafe Sub glTexCoord2d Lib "opengl32" (ByVal s As Double, ByVal t As Double)
    Public Declare PtrSafe Sub glPushMatrix Lib "opengl32" ()
    Public Declare PtrSafe Sub glPopMatrix Lib "opengl32" ()
    Public Declare PtrSafe Sub glBindTexture Lib "opengl32" (ByVal target As Long, ByVal texture As Long)
    Public Declare PtrSafe Sub glTexImage2D Lib "opengl32" (ByVal target As Long, ByVal level As Long, ByVal internalformat As Long, ByVal Width As Long, ByVal Height As Long, ByVal border As Long, ByVal format As Long, ByVal pixelType As Long, pixels As Any)
    Public Declare PtrSafe Sub glTexParameteri Lib "opengl32" (ByVal target As Long, ByVal pname As Long, ByVal param As Long)
    Public Declare PtrSafe Sub glTexGeni Lib "opengl32" (ByVal coord As Long, ByVal pname As Long, ByVal param As Long)
    Public Declare PtrSafe Sub glGenTextures Lib "opengl32" (ByVal n As Long, ByRef textures As Long)
    Public Declare PtrSafe Sub glDeleteTextures Lib "opengl32" (ByVal n As Long, ByRef textures As Long)
    Public Declare PtrSafe Sub glShadeModel Lib "opengl32" (ByVal mode As Long)
    Public Declare PtrSafe Sub glBlendFunc Lib "opengl32" (ByVal sfactor As Long, ByVal dfactor As Long)
    Public Declare PtrSafe Sub glCullFace Lib "opengl32" (ByVal mode As Long)
    ' GDI+
    Public Declare PtrSafe Function GdiplusStartup Lib "gdiplus" (token As LongPtr, inputbuf As GDIPlusStartupInput, output As Any) As Long
    Public Declare PtrSafe Function GdiplusShutdown Lib "gdiplus" (ByVal token As LongPtr) As Long
    Public Declare PtrSafe Function GdipCreateBitmapFromScan0 Lib "gdiplus" (ByVal Width As Long, ByVal Height As Long, ByVal Stride As Long, ByVal PixelFormat As Long, Scan0 As Any, bitmap As LongPtr) As Long
    Public Declare PtrSafe Function GdipLoadImageFromFile Lib "gdiplus" (ByVal filename As LongPtr, image As LongPtr) As Long
    Public Declare PtrSafe Function GdipGetImageWidth Lib "gdiplus" (ByVal image As LongPtr, Width As Long) As Long
    Public Declare PtrSafe Function GdipGetImageHeight Lib "gdiplus" (ByVal image As LongPtr, Height As Long) As Long
    Public Declare PtrSafe Function GdipGetImageGraphicsContext Lib "gdiplus" (ByVal image As LongPtr, graphics As LongPtr) As Long
    Public Declare PtrSafe Function GdipDeleteGraphics Lib "gdiplus" (ByVal graphics As LongPtr) As Long
    Public Declare PtrSafe Function GdipDisposeImage Lib "gdiplus" (ByVal image As LongPtr) As Long
    Public Declare PtrSafe Function GdipDrawString Lib "gdiplus" (ByVal graphics As LongPtr, ByVal str As LongPtr, ByVal length As Long, ByVal font As LongPtr, layoutRect As RECTF, ByVal stringFormat As LongPtr, ByVal brush As LongPtr) As Long
    Public Declare PtrSafe Function GdipCreateFont Lib "gdiplus" (ByVal fontFamily As LongPtr, ByVal emSize As Single, ByVal style As Long, ByVal unit As Long, font As LongPtr) As Long
    Public Declare PtrSafe Function GdipCreateFontFamilyFromName Lib "gdiplus" (ByVal name As LongPtr, ByVal fontCollection As LongPtr, fontFamily As LongPtr) As Long
    Public Declare PtrSafe Function GdipDeleteFontFamily Lib "gdiplus" (ByVal fontFamily As LongPtr) As Long
    Public Declare PtrSafe Function GdipDeleteFont Lib "gdiplus" (ByVal font As LongPtr) As Long
    Public Declare PtrSafe Function GdipCreateSolidFill Lib "gdiplus" (ByVal color As Long, brush As LongPtr) As Long
    Public Declare PtrSafe Function GdipDeleteBrush Lib "gdiplus" (ByVal brush As LongPtr) As Long
    Public Declare PtrSafe Function GdipGraphicsClear Lib "gdiplus" (ByVal graphics As LongPtr, ByVal color As Long) As Long
    Public Declare PtrSafe Function GdipBitmapLockBits Lib "gdiplus" (ByVal bitmap As LongPtr, rect As Any, ByVal flags As Long, ByVal PixelFormat As Long, lockedBitmapData As BitmapData) As Long
    Public Declare PtrSafe Function GdipBitmapUnlockBits Lib "gdiplus" (ByVal bitmap As LongPtr, lockedBitmapData As BitmapData) As Long
    ' Performance
    Public Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
    Public Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
#End If

' Constants
Public Const SW_SHOW = 5: Public Const IDC_ARROW = 32512&: Public Const IDC_HAND = 32649&: Public Const CS_HREDRAW = &H2: Public Const CS_VREDRAW = &H1: Public Const CS_OWNDC = &H20
Public Const WS_OVERLAPPEDWINDOW = &HCF0000: Public Const WS_CLIPSIBLINGS = &H4000000: Public Const WS_CLIPCHILDREN = &H2000000: Public Const CW_USEDEFAULT = &H80000000
Public Const WM_CLOSE = &H10: Public Const WM_DESTROY = &H2: Public Const PM_REMOVE = &H1: Public Const WM_LBUTTONDOWN = &H201: Public Const WM_LBUTTONUP = &H202: Public Const WM_MOUSEMOVE = &H200: Public Const WM_MOUSEWHEEL = &H20A
Public Const PFD_DRAW_TO_WINDOW = &H4: Public Const PFD_SUPPORT_OPENGL = &H20: Public Const PFD_DOUBLEBUFFER = &H1: Public Const PFD_TYPE_RGBA = 0
Public Const GL_COLOR_BUFFER_BIT = &H4000&: Public Const GL_DEPTH_BUFFER_BIT = &H100&: Public Const GL_DEPTH_TEST = &HB71&: Public Const GL_QUADS = &H7: Public Const GL_TRIANGLE_STRIP = &H5&
Public Const GL_PROJECTION = &H1701&: Public Const GL_MODELVIEW = &H1700&: Public Const GL_TEXTURE_2D = &HDE1&: Public Const GL_SMOOTH = &H1D01&
Public Const GL_LIGHTING = &HB50&: Public Const GL_LIGHT0 = &H4000&: Public Const GL_FRONT = &H404&: Public Const GL_CULL_FACE = &HB44&: Public Const GL_BLEND = &HBE2&: Public Const GL_SRC_ALPHA = &H302&: Public Const GL_ONE_MINUS_SRC_ALPHA = &H303&
Public Const GL_RGBA = &H1908&: Public Const GL_BGRA = &H80E1&: Public Const GL_UNSIGNED_BYTE = &H1401&: Public Const GL_LINEAR = &H2601&: Public Const GL_TEXTURE_MIN_FILTER = &H2801&: Public Const GL_TEXTURE_MAG_FILTER = &H2800&
Public Const GL_TEXTURE_GEN_S = &H1900: Public Const GL_TEXTURE_GEN_T = &H1901: Public Const GL_TEXTURE_GEN_MODE = &H2500: Public Const GL_SPHERE_MAP = &H2402: Public Const GL_S = &H2000: Public Const GL_T = &H2001
Public Const PixelFormat32bppARGB = &H26200A: Public Const ImageLockModeRead = &H1

Public Type PIXELFORMATDESCRIPTOR: nSize As Integer: nVersion As Integer: dwFlags As Long: iPixelType As Byte: cColorBits As Byte: cRedBits As Byte: cRedShift As Byte: cGreenBits As Byte: cGreenShift As Byte: cBlueBits As Byte: cBlueShift As Byte: cAlphaBits As Byte: cAlphaShift As Byte: cAccumBits As Byte: cAccumRedBits As Byte: cAccumGreenBits As Byte: cAccumBlueBits As Byte: cAccumAlphaBits As Byte: cDepthBits As Byte: cStencilBits As Byte: cAuxBuffers As Byte: iLayerType As Byte: bReserved As Byte: dwLayerMask As Long: dwVisibleMask As Long: dwDamageMask As Long: End Type
Public Type WNDCLASSEX: cbSize As Long: style As Long: lpfnWndProc As LongPtr: cbClsExtra As Long: cbWndExtra As Long: hInstance As LongPtr: hIcon As LongPtr: hCursor As LongPtr: hbrBackground As LongPtr: lpszMenuName As String: lpszClassName As String: hIconSm As LongPtr: End Type
Public Type POINTAPI: x As Long: y As Long: End Type
Public Type msg: hWnd As LongPtr: message As Long: wParam As LongPtr: lParam As LongPtr: time As Long: pt As POINTAPI: End Type
Public Type GDIPlusStartupInput: GdiplusVersion As Long: DebugEventCallback As LongPtr: SuppressBackgroundThread As Long: SuppressExternalCodecs As Long: End Type
Public Type RECTF: Left As Single: Top As Single: Right As Single: Bottom As Single: End Type
Public Type BitmapData: Width As Long: Height As Long: Stride As Long: PixelFormat As Long: Scan0 As LongPtr: Reserved As LongPtr: End Type

Public Function GlobalWndProc(ByVal hWnd As LongPtr, ByVal uMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
    If Not g_OpenGLWindow Is Nothing Then
        GlobalWndProc = g_OpenGLWindow.HandleMessage(hWnd, uMsg, wParam, lParam)
    Else
        GlobalWndProc = DefWindowProc(hWnd, uMsg, wParam, lParam)
    End If
End Function

Public Function SetWndProc(ByVal pfn As LongPtr) As LongPtr: SetWndProc = pfn: End Function
