VBA OpenGL PhysX Engine v2.0
A high-performance 3D engine built entirely in VBA (Visual Basic for Applications). This project leverages the Windows API and GDI+ to host a hardware-accelerated OpenGL context directly within a standalone window, featuring real-time physics, interactive UI, and 360° panoramic environments.
<img width="1013" height="777" alt="data=l8728u5aribYrmj4XFsUK-3G9M2H0cb5obSF9enIRSTXw6E1Shp0goCx7hkltjBjS8T98_wSIWVDrK-TmypY4UtMGO4RDsz2K4cLSkNN5B3bCBr3NUoLqtrP4vZW2biAm6We11pgGAd19YG5Db_DIUaop_cqKqrVpMREGivVAtJeQSmKJ4B7e6wzy-hW4-yA6EiqbQ" src="https://github.com/user-attachments/assets/426ae7a1-b5f5-40ba-9d3c-21d3cbf0877e" />

🚀 Key Features
Real-Time Sphere Physics: Custom physics engine supporting sphere-to-sphere and sphere-to-wall collisions.

Anti-Jiggle Logic: Advanced static separation and friction (damping) calculations to ensure stable resting states.

Interactive 2D UI: An OpenGL-rendered overlay with responsive buttons for resetting physics and toggling modes.

360° Panoramic Skybox: Support for equirectangular textures with smooth mouse-look navigation.

Dynamic UV Mapping: Precise texture coordinates for spheres and cubes, moving away from simple auto-generation to detailed image wrapping.

Hardware Accelerated: Bypasses slow GDI drawing for raw GPU performance, throttled to a stable 60 FPS.

🛠 Technical Implementation
Physics Engine
The engine uses a Velocity-Verlet style approach for motion, combined with an impulse-based collision resolver.

Friction: Velocity is multiplied by a FRICTION constant (0.98) each frame to simulate air resistance.

Collision: Uses a unit-vector normal to calculate elastic bounces and physically separates overlapping objects to prevent high-frequency vibration ("jiggling").

UI Architecture
The UI is managed via a constant-driven coordinate system, allowing for easy adjustment of margins and spacing.

VBA
Private Const BTN_GAP As Long = 10 
Private Const BTN_HEIGHT As Long = 35
It utilizes a 2D Orthographic projection overlaying the 3D Perspective frustum.

Assets & Compatibility
Textures: Loads .png, .jpg, and .bmp via GDI+ into OpenGL Texture Objects.

Environment: Includes a "Hotspot" detection system that triggers events based on your viewing angle (Yaw/Pitch) within the 360° space.

📂 Installation & Setup
Download the .cls and .bas files.

Import them into any VBA-enabled host (Excel, Access, or Word).

Ensure you have the required textures (map.png, 360_bg.jpg) in your project directory.

Run InitDemo to launch the visualization window.

📜 Version History
v2.0 (Current): Added Physics, UV Texturing, and constant-based UI.

v1.0 (Legacy): Initial proof-of-concept OpenGL renderer (available in the Releases tab).

🤝 Contributing
Feel free to fork this repository and submit pull requests. I'm currently exploring 2D Popups and potential Video Texture integration for future releases.
