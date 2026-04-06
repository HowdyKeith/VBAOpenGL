# VBAOpenGL

**Run OpenGL visualizations directly from VBA code in a separate window.**

This library allows Excel (and other Office) VBA developers to create hardware-accelerated 3D graphics, animations, and visualizations without leaving the VBA environment.

## Features
- Create an OpenGL rendering context from VBA
- Simple wrapper functions for common OpenGL calls
- Hardware-accelerated rendering in a dedicated window
- Easy integration with existing Excel/VBA projects
- [Add more specific features as you implement them]

## Requirements
- Windows (OpenGL context creation relies on Win32 APIs)
- VBA6/VBA7 (Excel 2007+ recommended)
- OpenGL drivers (most modern GPUs support it)

## Installation
1. Download or clone this repository.

## Quick Start / Example

```vba
Sub OpenGLDemo()
    Dim glWindow As New OpenGLWindow
    
    glWindow.Create 800, 600, "VBA OpenGL Demo"
    
    ' Simple render loop example
    Do While glWindow.IsRunning
        glClear GL_COLOR_BUFFER_BIT Or GL_DEPTH_BUFFER_BIT
        
        ' Your drawing code here (triangles, cubes, etc.)
        DrawSomething
        
        glWindow.SwapBuffers
        DoEvents
    Loop
    
    glWindow.Destroy
End Sub
