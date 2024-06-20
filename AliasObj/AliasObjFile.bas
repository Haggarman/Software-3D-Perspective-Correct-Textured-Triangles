Option _Explicit
_Title "Alias Object File 33"
' 2024 Haggarman
'  V31 Mirror Reflective Surface
'  V30 Skybox
'  V26 Experiments with using half-angle instead of bounce reflection for specular.
'  V23 Specular gouraud.
'  V19 Re-introduce texture mapping, although you only get red brick for now.
'  V17 is where the directional lighting using vertex normals is functioning.
'
'  Press G to toggle lighting method, if vertex normals are available.
'
' Camera and matrix math code translated from the works of Javidx9 OneLoneCoder.
' Texel interpolation and triangle drawing code by me.
'
Declare Library ""
    ' grab a C99 function from math.h
    Function powf! (ByVal ba!, Byval ex!)
End Declare

Dim Shared DISP_IMAGE As Long
Dim Shared WORK_IMAGE As Long
Dim Shared Size_Screen_X As Integer, Size_Screen_Y As Integer
Dim Shared Size_Render_X As Integer, Size_Render_Y As Integer
Dim Actor_Count As Integer
Actor_Count = 1 ' keep at 1 for now
Dim Shared Camera_Start_Z
Dim Obj_File_Name As String


' MODIFY THESE if you want.
Obj_File_Name = "" ' "bunny.obj" "cube.obj" "teacup.obj" "spoonfix.obj"
Size_Screen_X = 1024
Size_Screen_Y = 768
Size_Render_X = Size_Screen_X \ 2 ' render size
Size_Render_Y = Size_Screen_Y \ 2
Camera_Start_Z = -6.0

DISP_IMAGE = _NewImage(Size_Screen_X, Size_Screen_Y, 32)
Screen DISP_IMAGE

WORK_IMAGE = _NewImage(Size_Render_X, Size_Render_Y, 32)
_DontBlend

Dim Shared Screen_Z_Buffer_MaxElement As Long
Screen_Z_Buffer_MaxElement = Size_Render_X * Size_Render_Y - 1
Dim Shared Screen_Z_Buffer(Screen_Z_Buffer_MaxElement) As Single

Type vec3d
    x As Single
    y As Single
    z As Single
End Type

Type vec4d
    x As Single
    y As Single
    z As Single
    w As Single
End Type

Type triangle
    x0 As Single
    y0 As Single
    z0 As Single
    x1 As Single
    y1 As Single
    z1 As Single
    x2 As Single
    y2 As Single
    z2 As Single

    u0 As Single
    v0 As Single
    u1 As Single
    v1 As Single
    u2 As Single
    v2 As Single

    texture As _Unsigned Long
    options As _Unsigned Long
    material As Long
    vni0 As Long ' vector normal index
    vni1 As Long
    vni2 As Long
End Type

Type skybox_triangle
    x0 As Single
    y0 As Single
    z0 As Single
    x1 As Single
    y1 As Single
    z1 As Single
    x2 As Single
    y2 As Single
    z2 As Single

    u0 As Single
    v0 As Single
    u1 As Single
    v1 As Single
    u2 As Single
    v2 As Single
    texture As _Unsigned Long
End Type

Type vertex5
    x As Single
    y As Single
    w As Single
    u As Single
    v As Single
End Type

Type vertex9
    x As Single
    y As Single
    w As Single
    u As Single
    v As Single
    r As Single
    g As Single
    b As Single
    a As Single ' alpha ranges from 0.0 to 1.0 for less conversion calculations
End Type

Type vertex_attribute5
    u As Single
    v As Single
    r As Single
    g As Single
    b As Single
End Type

Type vertex_attribute7
    u As Single
    v As Single
    r As Single
    g As Single
    b As Single
    s As Single
    t As Single
End Type

Type objectlist_type
    viewz As Single
    psn As vec3d
    first As Long
    last As Long
End Type

Type newmtl_type
    Kd_r As Single
    Kd_g As Single
    Kd_b As Single
    Ks_r As Single
    Ks_g As Single
    Ks_b As Single
    diaphaneity As Single ' translucency factor
    textName As String
End Type

' Projection Matrix
Dim Shared Frustum_Near As Single
Dim Frustum_Far As Single
Dim Frustum_FOV_deg As Single
Dim Frustum_Aspect_Ratio As Single
Dim Frustum_FOV_ratio As Single

Frustum_Near = 0.1
Frustum_Far = 1000.0
Frustum_FOV_deg = 60.0
Frustum_Aspect_Ratio = _Height / _Width
Frustum_FOV_ratio = 1.0 / Tan(_D2R(Frustum_FOV_deg * 0.5))

Dim matProj(3, 3) As Single
matProj(0, 0) = Frustum_Aspect_Ratio * Frustum_FOV_ratio
matProj(1, 1) = Frustum_FOV_ratio
matProj(2, 2) = Frustum_Far / (Frustum_Far - Frustum_Near)
matProj(2, 3) = 1.0
matProj(3, 2) = (-Frustum_Far * Frustum_Near) / (Frustum_Far - Frustum_Near)
matProj(3, 3) = 0.0

' Viewing area clipping
Dim Shared clip_min_y As Long, clip_max_y As Long
Dim Shared clip_min_x As Long, clip_max_x As Long
clip_min_y = 1
clip_max_y = Size_Render_Y - 2
clip_min_x = 1
clip_max_x = Size_Render_X - 1

' Fog
Dim Shared Fog_near As Single, Fog_far As Single, Fog_rate As Single
Dim Shared Fog_color As Long
Dim Shared Fog_R As Long, Fog_G As Long, Fog_B As Long
Fog_near = 9.0
Fog_far = 19.0
Fog_rate = 1.0 / (Fog_far - Fog_near)

Fog_color = _RGB32(111, 177, 233)
Fog_R = _Red(Fog_color)
Fog_G = _Green(Fog_color)
Fog_B = _Blue(Fog_color)

' Z Fight has to do with overdrawing on top of the same coplanar surface.
' If it is positive, a newer pixel at the same exact Z will always overdraw the older one.
Dim Shared Z_Fight_Bias
Z_Fight_Bias = 0 ' -0.001953125 / 32.0


' These T1 Texture characteristics are read later on during drawing.
Dim Shared T1_ImageHandle As Long
Dim Shared T1_width As Integer, T1_height As Integer
Dim Shared T1_width_MASK As Integer, T1_height_MASK As Integer
Dim Shared T1_mblock As _MEM
Dim Shared T1_Filter_Selection As Integer
Dim Shared T1_last_cache As _Unsigned Long
Dim Shared T1_options As _Unsigned Long
Const T1_option_clamp_width = 1
Const T1_option_clamp_height = 2
Const T1_option_no_Z_write = 4
Const T1_option_no_T1 = 65536

' Later optimization in ReadTexel requires these to be powers of 2.
' That means: 2,4,8,16,32,64,128,256...
T1_width = 16: T1_width_MASK = T1_width - 1
T1_height = 16: T1_height_MASK = T1_height - 1

' Load Texture1 Array from Data
Dim Shared Texture1(T1_width_MASK, T1_height_MASK) As _Unsigned Long
Dim dvalue As _Unsigned Long
Restore Texture1Data
Dim row As Integer, col As Integer
For row = 0 To T1_height_MASK
    For col = 0 To T1_width_MASK
        Read dvalue
        Texture1(col, row) = dvalue
        'PSet (col, row), dvalue
    Next col
Next row

' Load Skybox
'     +---+
'     | 2 |
' +---+---+---+---+
' | 1 | 4 | 0 | 5 |
' +---+---+---+---+
'     | 3 |
'     +---+

Dim Shared SkyBoxRef(5) As Long
SkyBoxRef(0) = _LoadImage("SkyBoxRight.png", 32)
SkyBoxRef(1) = _LoadImage("SkyBoxLeft.png", 32)
SkyBoxRef(2) = _LoadImage("SkyBoxTop.png", 32)
SkyBoxRef(3) = _LoadImage("SkyBoxBottom.png", 32)
SkyBoxRef(4) = _LoadImage("SkyBoxFront.png", 32)
SkyBoxRef(5) = _LoadImage("SkyBoxBack.png", 32)

' Error _LoadImage returns -1 as an invalid handle if it cannot load the image.
Dim refIndex As Integer
For refIndex = 0 To 5
    If SkyBoxRef(refIndex) = -1 Then
        Print "Could not load texture file for skybox face: "; refIndex
        End
    End If
Next refIndex

Dim A As Integer
Dim Shared Sky_Last_Element As Integer
Sky_Last_Element = 11
Dim sky(Sky_Last_Element) As skybox_triangle
Restore SKYBOX
For A = 0 To Sky_Last_Element
    Read sky(A).x0
    Read sky(A).y0
    Read sky(A).z0

    Read sky(A).x1
    Read sky(A).y1
    Read sky(A).z1

    Read sky(A).x2
    Read sky(A).y2
    Read sky(A).z2

    Read sky(A).u0
    Read sky(A).v0
    Read sky(A).u1
    Read sky(A).v1
    Read sky(A).u2
    Read sky(A).v2

    Read sky(A).texture
Next A

' Load Mesh
While Obj_File_Name = ""
    Obj_File_Name = _OpenFileDialog$("Load Alias Object File", , "*.OBJ|*.obj")
    If Obj_File_Name = "" Then End
Wend

Dim Shared Objects_Last_Element As Integer
Objects_Last_Element = Actor_Count
Dim Objects(Objects_Last_Element) As objectlist_type
' index 0 will be invisible

Dim Shared Mesh_Last_Element As Long ' number of triangles, 1 is the minimum
Dim Shared Vertex_Count As Long ' 3 is the minimum
Dim Shared TextureCoord_Count As Long ' can be 0, textures are optional
Dim Shared Vtx_Normals_Count As Long ' can be 0, vertex normals are optional
Dim Shared Material_File_Count As Long
Dim Shared Material_File_Name As String
Dim Shared Materials_Count As Long

PrescanMesh Obj_File_Name, Mesh_Last_Element, Vertex_Count, TextureCoord_Count, Vtx_Normals_Count, Material_File_Count, Material_File_Name
If Mesh_Last_Element = 0 Then
    Color _RGB(249, 161, 50)
    Print "Error Loading Object File "; Obj_File_Name
    End
End If

If Material_File_Count >= 1 Then
    PrescanMaterialFile Material_File_Name, Materials_Count
    If Materials_Count = 0 Then
        Color _RGB(249, 161, 50)
        Print "Error Loading Material File "; Material_File_Name
        End
    End If
End If

' always create at least one material. 0 just creates a single element index 0.
Dim Shared Materials(Materials_Count) As newmtl_type
' sensible defaults to at least show something.
Materials(0).diaphaneity = 1.0
Materials(0).Kd_r = 0.5: Materials(0).Kd_g = 0.5: Materials(0).Kd_b = 0.5
Materials(0).Ks_r = 0.25: Materials(0).Ks_g = 0.25: Materials(0).Ks_b = 0.25

If Material_File_Count >= 1 Then
    LoadMaterialFile Material_File_Name, Materials(), Materials_Count
End If

' create and start reading in the triangle mesh
Dim Shared mesh(Mesh_Last_Element) As triangle

' 6-1-2024
Dim Shared vtxnorms(Vtx_Normals_Count) As vec3d

Dim actor As Integer
Dim tri As Long
tri = 0
For actor = 1 To Actor_Count
    Objects(actor).first = tri + 1
    LoadMesh Obj_File_Name, mesh(), tri, Vertex_Count, TextureCoord_Count, Materials(), vtxnorms()
    Objects(actor).last = tri
Next actor

' Here are the 3D math and projection vars

' Rotation
Dim matRotZ(3, 3) As Single
Dim matRotX(3, 3) As Single

Dim point0 As vec3d
Dim point1 As vec3d
Dim point2 As vec3d

Dim pointRotZ0 As vec3d
Dim pointRotZ1 As vec3d
Dim pointRotZ2 As vec3d

Dim pointRotZX0 As vec3d
Dim pointRotZX1 As vec3d
Dim pointRotZX2 As vec3d

' Translation (as in offset)
Dim pointWorld0 As vec3d
Dim pointWorld1 As vec3d
Dim pointWorld2 As vec3d

' View Space 2-10-2023
Dim matView(3, 3) As Single
Dim pointView0 As vec3d
Dim pointView1 As vec3d
Dim pointView2 As vec3d
Dim pointView3 As vec3d ' extra clipped tri

' Near frustum clipping 2-27-2023
Dim Shared vatr0 As vertex_attribute7
Dim Shared vatr1 As vertex_attribute7
Dim Shared vatr2 As vertex_attribute7
Dim Shared vatr3 As vertex_attribute7

' Projection
Dim pointProj0 As vec4d ' added w
Dim pointProj1 As vec4d
Dim pointProj2 As vec4d
Dim pointProj3 As vec4d ' extra clipped tri

' Surface Normal Calculation
' Part 2
Dim tri_normal As vec3d
' 6-1-2024
Dim vertex_normal_A As vec3d
Dim vertex_normal_B As vec3d
Dim vertex_normal_C As vec3d
Dim object_vtx_normals(Vtx_Normals_Count) As vec3d

' Part 2-2
Dim vCameraPsn As vec3d ' location of camera in world space
vCameraPsn.x = 0.0
vCameraPsn.y = 0.0
vCameraPsn.z = Camera_Start_Z

Dim cameraRay0 As vec3d
Dim dotProductCam As Single

' View Space 2-10-2023
Dim fPitch As Single ' FPS Camera rotation in YZ plane (X)
Dim fYaw As Single ' FPS Camera rotation in XZ plane (Y)
Dim fRoll As Single
Dim matCameraRot(3, 3) As Single

Dim vCameraHomeFwd As vec3d ' Home angle orientation is facing down the Z line.
vCameraHomeFwd.x = 0.0: vCameraHomeFwd.y = 0.0: vCameraHomeFwd.z = 1.0

Dim vCameraTripod As vec3d ' Home angle orientation of which way is up.
' You could simulate tipping over the camera tripod with something other than y=1, and it will gimbal oddly.
vCameraTripod.x = 0.0: vCameraTripod.y = 1.0: vCameraTripod.z = 0.0

Dim vLookPitch As vec3d
Dim vLookDir As vec3d
Dim vCameraTarget As vec3d
Dim matCamera(3, 3) As Single


' Directional light 1-17-2023
Dim vLightDir As vec3d
' Put the light source where the camera starts so you don't go insane when trying to get the specular vectors correct.
vLightDir.x = 0.0
vLightDir.y = 0.0 ' +Y is up
vLightDir.z = Camera_Start_Z
Vector3_Normalize vLightDir
Dim Shared Light_Directional As Single
Dim Shared Light_AmbientVal As Single
Light_AmbientVal = 0.2

' Directional lighting using vertex normals
Dim light_directional_A As Single
Dim light_directional_B As Single
Dim light_directional_C As Single
Dim face_light_r As Single
Dim face_light_g As Single
Dim face_light_b As Single

' Specular lighting
Dim cameraRay1 As vec3d
Dim cameraRay2 As vec3d
Dim vLightSpec As vec3d
Dim vHalfAngle As vec3d
Dim reflectLightDir0 As vec3d
Dim reflectLightDir1 As vec3d
Dim reflectLightDir2 As vec3d
Dim light_specular_A As Single
Dim light_specular_B As Single
Dim light_specular_C As Single
Dim Default_Specular_Power As Single
Dim Default_Specular_HA_Power As Single
Default_Specular_Power = 16.0 ' when bouncing the ray.
' Half angle specular power needs to obey conservation of energy.
Default_Specular_HA_Power = 4.0 * Default_Specular_Power

' Screen Scaling
Dim halfWidth As Single
Dim halfHeight As Single
halfWidth = Size_Render_X / 2
halfHeight = Size_Render_Y / 2

' Projected Screen Coordinate List
Dim SX0 As Single, SY0 As Single
Dim SX1 As Single, SY1 As Single
Dim SX2 As Single, SY2 As Single
Dim SX3 As Single, SY3 As Single

Dim vertexA As vertex9
Dim vertexB As vertex9
Dim vertexC As vertex9
Dim Shared envMapReflectionRayA As vec3d
Dim Shared envMapReflectionRayB As vec3d
Dim Shared envMapReflectionRayC As vec3d

' This is so that the cube object animates by rotating
Dim spinAngleDegZ As Single
Dim spinAngleDegX As Single
spinAngleDegZ = 0.0
spinAngleDegX = 0.0

' code execution time
Dim start_ms As Double
Dim render_ms As Double

' physics framerate
Dim frametime_fullframe_ms As Double
Dim frametime_fullframethreshold_ms As Double
Dim frametimestamp_now_ms As Double
Dim frametimestamp_prior_ms As Double
Dim frametimestamp_delta_ms As Double
Dim frame_advance As Integer

' Main loop stuff
Dim renderPass As Integer
Dim transparencyFactor As Single
Dim KeyNow As String
Dim ExitCode As Integer
Dim triCount As Integer
Dim Animate_Spin As Integer
Dim Shared Dither_Selection As Integer
Dim Gouraud_Shading_Selection As Integer
Dim vMove_Player_Forward As vec3d
Dim thisMaterial As newmtl_type


main:
$Checking:Off
ExitCode = 0
Animate_Spin = -1
T1_Filter_Selection = 2
Dither_Selection = 0
Gouraud_Shading_Selection = 4
actor = 1

fPitch = 0.0
fYaw = 0.0
fRoll = 0.0

frametime_fullframe_ms = 1 / 60.0
frametime_fullframethreshold_ms = 1 / 61.0
frametimestamp_prior_ms = Timer(.001)
frametimestamp_delta_ms = frametime_fullframe_ms
frame_advance = 0

Do
    If Animate_Spin Then
        spinAngleDegZ = spinAngleDegZ + frame_advance * 0.355
        spinAngleDegX = spinAngleDegX - frame_advance * 0.23
        'fYaw = fYaw + 1
    End If

    ' Set up rotation matrices
    ' _D2R is just a built-in degrees to radians conversion

    ' Rotation Z
    matRotZ(0, 0) = Cos(_D2R(spinAngleDegZ))
    matRotZ(0, 1) = Sin(_D2R(spinAngleDegZ))
    matRotZ(1, 0) = -Sin(_D2R(spinAngleDegZ))
    matRotZ(1, 1) = Cos(_D2R(spinAngleDegZ))
    matRotZ(2, 2) = 1
    matRotZ(3, 3) = 1

    ' Rotation X
    matRotX(0, 0) = 1
    matRotX(1, 1) = Cos(_D2R(spinAngleDegX))
    matRotX(1, 2) = -Sin(_D2R(spinAngleDegX)) ' flip
    matRotX(2, 1) = Sin(_D2R(spinAngleDegX)) ' flip
    matRotX(2, 2) = Cos(_D2R(spinAngleDegX))
    matRotX(3, 3) = 1


    ' Create "Point At" Matrix for camera

    ' the neck tilts up and down first
    Matrix4_MakeRotation_X fPitch, matCameraRot()
    Multiply_Vector3_Matrix4 vCameraHomeFwd, matCameraRot(), vLookPitch

    ' then you spin around in place
    Matrix4_MakeRotation_Y fYaw, matCameraRot()
    Multiply_Vector3_Matrix4 vLookPitch, matCameraRot(), vLookDir

    ' Add to camera position to chase a dangling carrot so to speak
    Vector3_Add vCameraPsn, vLookDir, vCameraTarget

    Matrix4_PointAt vCameraPsn, vCameraTarget, vCameraTripod, matCamera()

    ' Make view matrix from Camera
    Matrix4_QuickInverse matCamera(), matView()

    start_ms = Timer(.001)

    ' Due to Skybox being drawn always, do not need to Clear Screen
    _Dest WORK_IMAGE
    'Cls , Fog_color
    _Source WORK_IMAGE

    ' Clear Z-Buffer
    ' This is a qbasic only optimization. it sets the array to zero. it saves 10 ms.
    ReDim Screen_Z_Buffer(Screen_Z_Buffer_MaxElement)

    ' Draw Skybox 2-28-2023
    For A = 0 To Sky_Last_Element
        point0.x = sky(A).x0
        point0.y = sky(A).y0
        point0.z = sky(A).z0

        point1.x = sky(A).x1
        point1.y = sky(A).y1
        point1.z = sky(A).z1

        point2.x = sky(A).x2
        point2.y = sky(A).y2
        point2.z = sky(A).z2

        ' Follow the camera coordinate position (slide)
        ' Skybox is like putting your head inside a floating box that travels with you, but never rotates.
        ' You rotate your head inside the skybox as you look around.
        Vector3_Add point0, vCameraPsn, pointWorld0
        Vector3_Add point1, vCameraPsn, pointWorld1
        Vector3_Add point2, vCameraPsn, pointWorld2

        ' Part 2 (Triangle Surface Normal Calculation)
        CalcSurfaceNormal_3Point pointWorld0, pointWorld1, pointWorld2, tri_normal

        ' The dot product to this skybox surface is just the way you are facing.
        ' The surface completely behind you is going to get later removed with NearClip.
        'Vector3_Delta vCameraPsn, pointWorld0, cameraRay
        'dotProductCam = Vector3_DotProduct!(tri_normal, cameraRay)
        'If dotProductCam > 0.0 Then

        ' Convert World Space --> View Space
        Multiply_Vector3_Matrix4 pointWorld0, matView(), pointView0
        Multiply_Vector3_Matrix4 pointWorld1, matView(), pointView1
        Multiply_Vector3_Matrix4 pointWorld2, matView(), pointView2

        ' Load up attribute lists here because NearClip will interpolate those too.
        vatr0.u = sky(A).u0: vatr0.v = sky(A).v0
        vatr1.u = sky(A).u1: vatr1.v = sky(A).v1
        vatr2.u = sky(A).u2: vatr2.v = sky(A).v2

        ' Clip more often than not in this example
        NearClip pointView0, pointView1, pointView2, pointView3, vatr0, vatr1, vatr2, vatr3, triCount
        If triCount > 0 Then
            ' Project triangles from 3D -----------------> 2D
            ProjectMatrixVector4 pointView0, matProj(), pointProj0
            ProjectMatrixVector4 pointView1, matProj(), pointProj1
            ProjectMatrixVector4 pointView2, matProj(), pointProj2

            ' Slide to center, then Scale into viewport
            SX0 = (pointProj0.x + 1) * halfWidth
            SY0 = (pointProj0.y + 1) * halfHeight

            SX1 = (pointProj1.x + 1) * halfWidth
            SY1 = (pointProj1.y + 1) * halfHeight

            SX2 = (pointProj2.x + 1) * halfWidth
            SY2 = (pointProj2.y + 1) * halfHeight

            ' Load Vertex List for Single Textured triangle
            vertexA.x = SX0
            vertexA.y = SY0
            vertexA.w = pointProj0.w ' depth
            vertexA.u = vatr0.u * pointProj0.w
            vertexA.v = vatr0.v * pointProj0.w

            vertexB.x = SX1
            vertexB.y = SY1
            vertexB.w = pointProj1.w ' depth
            vertexB.u = vatr1.u * pointProj1.w
            vertexB.v = vatr1.v * pointProj1.w

            vertexC.x = SX2
            vertexC.y = SY2
            vertexC.w = pointProj2.w ' depth
            vertexC.u = vatr2.u * pointProj2.w
            vertexC.v = vatr2.v * pointProj2.w

            ' No Directional light

            ' Fill in Texture 1 data
            T1_ImageHandle = SkyBoxRef(sky(A).texture)
            T1_mblock = _MemImage(T1_ImageHandle)
            T1_width = _Width(T1_ImageHandle): T1_width_MASK = T1_width - 1
            T1_height = _Height(T1_ImageHandle): T1_height_MASK = T1_height - 1

            TexturedNonlitTriangle vertexA, vertexB, vertexC

            ' Wireframe triangle
            'Line (SX0, SY0)-(SX1, SY1), _RGB32(128, 0, 128)
            'Line (SX1, SY1)-(SX2, SY2), _RGB32(128, 0, 128)
            'Line (SX2, SY2)-(SX0, SY0), _RGB32(128, 0, 128)
        End If
        If triCount = 2 Then

            ProjectMatrixVector4 pointView3, matProj(), pointProj3
            SX3 = (pointProj3.x + 1) * halfWidth
            SY3 = (pointProj3.y + 1) * halfHeight

            ' Reload Vertex List for Textured triangle
            vertexA.x = SX0
            vertexA.y = SY0
            vertexA.w = pointProj0.w ' depth
            vertexA.u = vatr0.u * pointProj0.w
            vertexA.v = vatr0.v * pointProj0.w

            vertexB.x = SX2
            vertexB.y = SY2
            vertexB.w = pointProj2.w ' depth
            vertexB.u = vatr2.u * pointProj2.w
            vertexB.v = vatr2.v * pointProj2.w

            vertexC.x = SX3
            vertexC.y = SY3
            vertexC.w = pointProj3.w ' depth
            vertexC.u = vatr3.u * pointProj3.w
            vertexC.v = vatr3.v * pointProj3.w

            TexturedNonlitTriangle vertexA, vertexB, vertexC

            ' Wireframe triangle
            'Line (SX0, SY0)-(SX2, SY2), _RGB32(0, 128, 128)
            'Line (SX2, SY2)-(SX3, SY3), _RGB32(0, 128, 128)
            'Line (SX3, SY3)-(SX0, SY0), _RGB32(0, 128, 128)
        End If
    Next A

    ' It may be faster to pre-rotate the vertex normals.
    For tri = 1 To Vtx_Normals_Count
        ' Vertex normals can be rotated around their origin and still retain their effectiveness.
        'object_vtx_normals(tri) = vtxnorms(tri)

        ' Rotate in Z-Axis
        Multiply_Vector3_Matrix4 vtxnorms(tri), matRotZ(), pointRotZ0
        ' Rotate in X-Axis
        Multiply_Vector3_Matrix4 pointRotZ0, matRotX(), object_vtx_normals(tri)
    Next tri

    ' Draw Triangles
    For renderPass = 0 To 1
        For tri = Objects(actor).first To Objects(actor).last
            transparencyFactor = Materials(mesh(tri).material).diaphaneity
            If ((renderPass = 0) And (transparencyFactor < 1.0)) Or ((renderPass = 1) And (transparencyFactor = 1.0)) Then GoTo Lbl_Skip_tri

            point0.x = mesh(tri).x0
            point0.y = mesh(tri).y0
            point0.z = mesh(tri).z0

            point1.x = mesh(tri).x1
            point1.y = mesh(tri).y1
            point1.z = mesh(tri).z1

            point2.x = mesh(tri).x2
            point2.y = mesh(tri).y2
            point2.z = mesh(tri).z2

            ' Rotate in Z-Axis
            Multiply_Vector3_Matrix4 point0, matRotZ(), pointRotZ0
            Multiply_Vector3_Matrix4 point1, matRotZ(), pointRotZ1
            Multiply_Vector3_Matrix4 point2, matRotZ(), pointRotZ2

            ' Rotate in X-Axis
            Multiply_Vector3_Matrix4 pointRotZ0, matRotX(), pointRotZX0
            Multiply_Vector3_Matrix4 pointRotZ1, matRotX(), pointRotZX1
            Multiply_Vector3_Matrix4 pointRotZ2, matRotX(), pointRotZX2

            ' Offset into the screen
            pointWorld0 = pointRotZX0
            pointWorld1 = pointRotZX1
            pointWorld2 = pointRotZX2


            ' Part 2 (Triangle Surface Normal Calculation)
            CalcSurfaceNormal_3Point pointWorld0, pointWorld1, pointWorld2, tri_normal

            Vector3_Delta vCameraPsn, pointWorld0, cameraRay0 ' be careful, this is not normalized.
            dotProductCam = Vector3_DotProduct!(tri_normal, cameraRay0) ' only interested here in the sign (positive)

            If dotProductCam > 0.0 Then
                ' Convert World Space --> View Space
                Multiply_Vector3_Matrix4 pointWorld0, matView(), pointView0
                Multiply_Vector3_Matrix4 pointWorld1, matView(), pointView1
                Multiply_Vector3_Matrix4 pointWorld2, matView(), pointView2

                ' Skip if any Z is too close
                If (pointView0.z < Frustum_Near) Or (pointView1.z < Frustum_Near) Or (pointView2.z < Frustum_Near) Then
                    GoTo Lbl_Skip_tri
                End If

                ' Project triangles from 3D -----------------> 2D
                ProjectMatrixVector4 pointView0, matProj(), pointProj0
                ProjectMatrixVector4 pointView1, matProj(), pointProj1
                ProjectMatrixVector4 pointView2, matProj(), pointProj2

                ' Early scissor reject
                If pointProj0.x > 1.0 And pointProj1.x > 1.0 And pointProj2.x > 1.0 Then GoTo Lbl_Skip_tri
                If pointProj0.x < -1.0 And pointProj1.x < -1.0 And pointProj2.x < -1.0 Then GoTo Lbl_Skip_tri
                If pointProj0.y > 1.0 And pointProj1.y > 1.0 And pointProj2.y > 1.0 Then GoTo Lbl_Skip_tri
                If pointProj0.y < -1.0 And pointProj1.y < -1.0 And pointProj2.y < -1.0 Then GoTo Lbl_Skip_tri

                ' Slide to center, then Scale into viewport
                SX0 = (pointProj0.x + 1) * halfWidth
                SY0 = (pointProj0.y + 1) * halfHeight

                SX1 = (pointProj1.x + 1) * halfWidth
                SY1 = (pointProj1.y + 1) * halfHeight

                SX2 = (pointProj2.x + 1) * halfWidth
                SY2 = (pointProj2.y + 1) * halfHeight

                ' Load Vertex List for Textured triangle
                vertexA.x = SX0
                vertexA.y = SY0
                vertexA.w = pointProj0.w ' depth
                vertexA.u = mesh(tri).u0 * T1_width * pointProj0.w
                vertexA.v = mesh(tri).v0 * T1_height * pointProj0.w
                vertexA.a = transparencyFactor * pointProj0.w

                vertexB.x = SX1
                vertexB.y = SY1
                vertexB.w = pointProj1.w ' depth
                vertexB.u = mesh(tri).u1 * T1_width * pointProj1.w
                vertexB.v = mesh(tri).v1 * T1_height * pointProj1.w
                vertexB.a = transparencyFactor * pointProj1.w

                vertexC.x = SX2
                vertexC.y = SY2
                vertexC.w = pointProj2.w ' depth
                vertexC.u = mesh(tri).u2 * T1_width * pointProj2.w
                vertexC.v = mesh(tri).v2 * T1_height * pointProj2.w
                vertexC.a = transparencyFactor * pointProj2.w

                T1_options = mesh(tri).options

                If mesh(tri).texture = 0 Then
                    ' start with the diffuse color
                    T1_options = T1_options Or T1_option_no_T1
                End If

                thisMaterial = Materials(mesh(tri).material)

                Select Case Gouraud_Shading_Selection
                    Case 0:
                        ' Flat face shading

                        ' Directional light 1-17-2023
                        Light_Directional = Vector3_DotProduct!(tri_normal, vLightDir)
                        If Light_Directional < 0.0 Then Light_Directional = 0.0

                        ' Specular light
                        ' Instead of a normalized light position pointing out from origin, needs to be pointing inward towards the reflective surface.
                        vLightSpec.x = -vLightDir.x: vLightSpec.y = -vLightDir.y: vLightSpec.z = -vLightDir.z
                        Vector3_Reflect_unroll vLightSpec, tri_normal, reflectLightDir0
                        'Vector3_Normalize reflectLightDir0

                        ' cameraRay0 was already calculated for backface removal.
                        Vector3_Normalize cameraRay0
                        light_specular_A = Vector3_DotProduct!(reflectLightDir0, cameraRay0)
                        If light_specular_A > 0.0 Then
                            ' this power thing only works because it should range from 0..1 again.
                            ' so what it actually does is a higher power pushes the number towards 0 and makes the rolloff steeper.
                            light_specular_A = 3 * powf(light_specular_A, Default_Specular_Power)
                        Else
                            light_specular_A = 0.0
                        End If

                        face_light_r = 255.0 * (thisMaterial.Kd_r * Light_Directional + thisMaterial.Ks_r * light_specular_A + Light_AmbientVal)
                        face_light_g = 255.0 * (thisMaterial.Kd_g * Light_Directional + thisMaterial.Ks_g * light_specular_A + Light_AmbientVal)
                        face_light_b = 255.0 * (thisMaterial.Kd_b * Light_Directional + thisMaterial.Ks_b * light_specular_A + Light_AmbientVal)

                        vertexA.r = face_light_r
                        vertexA.g = face_light_g
                        vertexA.b = face_light_b

                        vertexB.r = face_light_r
                        vertexB.g = face_light_g
                        vertexB.b = face_light_b

                        vertexC.r = face_light_r
                        vertexC.g = face_light_g
                        vertexC.b = face_light_b

                    Case 1:
                        ' Smooth shading

                        ' 6-15-2024 pre-rotated normals
                        vertex_normal_A = object_vtx_normals(mesh(tri).vni0)
                        vertex_normal_B = object_vtx_normals(mesh(tri).vni1)
                        vertex_normal_C = object_vtx_normals(mesh(tri).vni2)

                        ' Directional light
                        light_directional_A = Vector3_DotProduct(vertex_normal_A, vLightDir)
                        light_directional_B = Vector3_DotProduct(vertex_normal_B, vLightDir)
                        light_directional_C = Vector3_DotProduct(vertex_normal_C, vLightDir)

                        If light_directional_A < 0.0 Then light_directional_A = 0.0
                        If light_directional_B < 0.0 Then light_directional_B = 0.0
                        If light_directional_C < 0.0 Then light_directional_C = 0.0

                        ' Specular at Vertex
                        ' Instead of a normalized light position pointing out from origin, needs to be pointing inward towards the reflective surface.
                        vLightSpec.x = -vLightDir.x: vLightSpec.y = -vLightDir.y: vLightSpec.z = -vLightDir.z

                        Vector3_Reflect_unroll vLightSpec, vertex_normal_A, reflectLightDir0
                        'Vector3_Normalize reflectLightDir0
                        Vector3_Reflect_unroll vLightSpec, vertex_normal_B, reflectLightDir1
                        'Vector3_Normalize reflectLightDir1
                        Vector3_Reflect_unroll vLightSpec, vertex_normal_C, reflectLightDir2
                        'Vector3_Normalize reflectLightDir2

                        ' A
                        ' cameraRay0 was already calculated for backface removal.
                        Vector3_Normalize cameraRay0
                        light_specular_A = Vector3_DotProduct!(reflectLightDir0, cameraRay0)
                        If light_specular_A > 0.0 Then
                            ' this power thing only works because it should range from 0..1 again.
                            ' so what it actually does is a higher power pushes the number towards 0 and makes the rolloff steeper.
                            light_specular_A = 3.0 * powf(light_specular_A, Default_Specular_Power)
                        Else
                            light_specular_A = 0.0
                        End If

                        ' B
                        Vector3_Delta vCameraPsn, pointWorld1, cameraRay1
                        Vector3_Normalize cameraRay1
                        light_specular_B = Vector3_DotProduct!(reflectLightDir1, cameraRay1)
                        If light_specular_B > 0.0 Then
                            light_specular_B = 3.0 * powf(light_specular_B, Default_Specular_Power)
                        Else
                            light_specular_B = 0.0
                        End If

                        ' C
                        Vector3_Delta vCameraPsn, pointWorld2, cameraRay2
                        Vector3_Normalize cameraRay2
                        light_specular_C = Vector3_DotProduct!(reflectLightDir2, cameraRay2)
                        If light_specular_C > 0.0 Then
                            light_specular_C = 3.0 * powf(light_specular_C, Default_Specular_Power)
                        Else
                            light_specular_C = 0.0
                        End If

                        vertexA.r = 255.0 * (thisMaterial.Kd_r * light_directional_A + thisMaterial.Ks_r * light_specular_A + Light_AmbientVal)
                        vertexA.g = 255.0 * (thisMaterial.Kd_g * light_directional_A + thisMaterial.Ks_g * light_specular_A + Light_AmbientVal)
                        vertexA.b = 255.0 * (thisMaterial.Kd_b * light_directional_A + thisMaterial.Ks_b * light_specular_A + Light_AmbientVal)

                        vertexB.r = 255.0 * (thisMaterial.Kd_r * light_directional_B + thisMaterial.Ks_r * light_specular_B + Light_AmbientVal)
                        vertexB.g = 255.0 * (thisMaterial.Kd_g * light_directional_B + thisMaterial.Ks_g * light_specular_B + Light_AmbientVal)
                        vertexB.b = 255.0 * (thisMaterial.Kd_b * light_directional_B + thisMaterial.Ks_b * light_specular_B + Light_AmbientVal)

                        vertexC.r = 255.0 * (thisMaterial.Kd_r * light_directional_C + thisMaterial.Ks_r * light_specular_C + Light_AmbientVal)
                        vertexC.g = 255.0 * (thisMaterial.Kd_g * light_directional_C + thisMaterial.Ks_g * light_specular_C + Light_AmbientVal)
                        vertexC.b = 255.0 * (thisMaterial.Kd_b * light_directional_C + thisMaterial.Ks_b * light_specular_C + Light_AmbientVal)

                    Case 2:
                        ' Smooth shading
                        ' Fake Half-Angle specular per vertex

                        ' 6-15-2024 pre-rotated normals
                        vertex_normal_A = object_vtx_normals(mesh(tri).vni0)
                        vertex_normal_B = object_vtx_normals(mesh(tri).vni1)
                        vertex_normal_C = object_vtx_normals(mesh(tri).vni2)

                        ' Directional light
                        light_directional_A = Vector3_DotProduct(vertex_normal_A, vLightDir)
                        light_directional_B = Vector3_DotProduct(vertex_normal_B, vLightDir)
                        light_directional_C = Vector3_DotProduct(vertex_normal_C, vLightDir)

                        If light_directional_A < 0.0 Then light_directional_A = 0.0
                        If light_directional_B < 0.0 Then light_directional_B = 0.0
                        If light_directional_C < 0.0 Then light_directional_C = 0.0

                        ' Specular light
                        ' A
                        ' cameraRay0 was already calculated for backface removal.
                        Vector3_Normalize cameraRay0
                        Vector3_Add cameraRay0, vLightDir, vHalfAngle
                        Vector3_Normalize vHalfAngle
                        light_specular_A = Vector3_DotProduct!(vHalfAngle, vertex_normal_A)
                        If light_specular_A > 0.0 Then
                            ' this power thing only works because it should range from 0..1 again.
                            ' so what it actually does is a higher power pushes the number towards 0 and makes the rolloff steeper.
                            light_specular_A = 3.0 * powf(light_specular_A, Default_Specular_HA_Power)
                        Else
                            light_specular_A = 0.0
                        End If

                        ' B
                        Vector3_Delta vCameraPsn, pointWorld1, cameraRay1
                        Vector3_Normalize cameraRay1
                        Vector3_Add cameraRay1, vLightDir, vHalfAngle
                        Vector3_Normalize vHalfAngle
                        light_specular_B = Vector3_DotProduct!(vHalfAngle, vertex_normal_B)
                        If light_specular_B > 0.0 Then
                            light_specular_B = 3.0 * powf(light_specular_B, Default_Specular_HA_Power)
                        Else
                            light_specular_B = 0.0
                        End If

                        ' C
                        Vector3_Delta vCameraPsn, pointWorld2, cameraRay2
                        Vector3_Normalize cameraRay2
                        Vector3_Add cameraRay2, vLightDir, vHalfAngle
                        Vector3_Normalize vHalfAngle
                        light_specular_C = Vector3_DotProduct!(vHalfAngle, vertex_normal_C)
                        If light_specular_C > 0.0 Then
                            light_specular_C = 3.0 * powf(light_specular_C, Default_Specular_HA_Power)
                        Else
                            light_specular_C = 0.0
                        End If

                        vertexA.r = 255.0 * (thisMaterial.Kd_r * light_directional_A + thisMaterial.Ks_r * light_specular_A + Light_AmbientVal)
                        vertexA.g = 255.0 * (thisMaterial.Kd_g * light_directional_A + thisMaterial.Ks_g * light_specular_A + Light_AmbientVal)
                        vertexA.b = 255.0 * (thisMaterial.Kd_b * light_directional_A + thisMaterial.Ks_b * light_specular_A + Light_AmbientVal)

                        vertexB.r = 255.0 * (thisMaterial.Kd_r * light_directional_B + thisMaterial.Ks_r * light_specular_B + Light_AmbientVal)
                        vertexB.g = 255.0 * (thisMaterial.Kd_g * light_directional_B + thisMaterial.Ks_g * light_specular_B + Light_AmbientVal)
                        vertexB.b = 255.0 * (thisMaterial.Kd_b * light_directional_B + thisMaterial.Ks_b * light_specular_B + Light_AmbientVal)

                        vertexC.r = 255.0 * (thisMaterial.Kd_r * light_directional_C + thisMaterial.Ks_r * light_specular_C + Light_AmbientVal)
                        vertexC.g = 255.0 * (thisMaterial.Kd_g * light_directional_C + thisMaterial.Ks_g * light_specular_C + Light_AmbientVal)
                        vertexC.b = 255.0 * (thisMaterial.Kd_b * light_directional_C + thisMaterial.Ks_b * light_specular_C + Light_AmbientVal)

                    Case 3:
                        ' Fake Half-Angle specular per face.

                        ' Directional light
                        Light_Directional = Vector3_DotProduct!(tri_normal, vLightDir)
                        If Light_Directional < 0.0 Then Light_Directional = 0.0

                        ' Specular light
                        Vector3_Normalize cameraRay0
                        Vector3_Add cameraRay0, vLightDir, vHalfAngle
                        Vector3_Normalize vHalfAngle

                        light_specular_A = Vector3_DotProduct!(vHalfAngle, tri_normal)
                        If light_specular_A > 0.0 Then
                            light_specular_A = 3 * powf(light_specular_A, Default_Specular_HA_Power)
                        Else
                            light_specular_A = 0.0
                        End If

                        face_light_r = 255.0 * (thisMaterial.Kd_r * Light_Directional + thisMaterial.Ks_r * light_specular_A + Light_AmbientVal)
                        face_light_g = 255.0 * (thisMaterial.Kd_g * Light_Directional + thisMaterial.Ks_g * light_specular_A + Light_AmbientVal)
                        face_light_b = 255.0 * (thisMaterial.Kd_b * Light_Directional + thisMaterial.Ks_b * light_specular_A + Light_AmbientVal)

                        vertexA.r = face_light_r
                        vertexA.g = face_light_g
                        vertexA.b = face_light_b

                        vertexB.r = face_light_r
                        vertexB.g = face_light_g
                        vertexB.b = face_light_b

                        vertexC.r = face_light_r
                        vertexC.g = face_light_g
                        vertexC.b = face_light_b

                    Case 4:
                        ' 6-15-2024 pre-rotated normals
                        vertex_normal_A = object_vtx_normals(mesh(tri).vni0)
                        vertex_normal_B = object_vtx_normals(mesh(tri).vni1)
                        vertex_normal_C = object_vtx_normals(mesh(tri).vni2)

                        Vector3_Delta pointWorld0, vCameraPsn, cameraRay0
                        Vector3_Reflect cameraRay0, vertex_normal_A, envMapReflectionRayA

                        Vector3_Delta pointWorld1, vCameraPsn, cameraRay1
                        Vector3_Reflect cameraRay1, vertex_normal_B, envMapReflectionRayB

                        Vector3_Delta pointWorld2, vCameraPsn, cameraRay2
                        Vector3_Reflect cameraRay2, vertex_normal_C, envMapReflectionRayC

                        vertexA.r = envMapReflectionRayA.x
                        vertexA.g = envMapReflectionRayA.y
                        vertexA.b = envMapReflectionRayA.z

                        vertexB.r = envMapReflectionRayB.x
                        vertexB.g = envMapReflectionRayB.y
                        vertexB.b = envMapReflectionRayB.z

                        vertexC.r = envMapReflectionRayC.x
                        vertexC.g = envMapReflectionRayC.y
                        vertexC.b = envMapReflectionRayC.z

                        ReflectionMapTriangle vertexA, vertexB, vertexC
                End Select

                If Gouraud_Shading_Selection < 4 Then
                    TexturedVertexColorAlphaTriangle vertexA, vertexB, vertexC
                End If

                ' Wireframe triangle
                'Line (SX0, SY0)-(SX1, SY1), _RGB32(128, 128, 128)
                'Line (SX1, SY1)-(SX2, SY2), _RGB32(128, 128, 128)
                'Line (SX2, SY2)-(SX0, SY0), _RGB32(128, 128, 128)
            End If

            Lbl_Skip_tri:
        Next tri
    Next renderPass



    render_ms = Timer(.001)

    _PutImage , WORK_IMAGE, DISP_IMAGE
    _Dest DISP_IMAGE
    Locate 1, 1
    Color _RGB32(177, 227, 255)
    Print Using "render time #.###"; render_ms - start_ms
    Color _RGB32(249, 244, 17)
    Print "ESC to exit. ";
    Color _RGB32(233)
    Print "Arrow Keys Move."

    If Animate_Spin Then
        Print "Press S to Stop Spin"
    Else
        Print "Press S to Start Spin"
    End If

    Print "Press G for Lighting: ";
    Select Case Gouraud_Shading_Selection
        Case 0
            Print "Flat using face normals"
        Case 1
            Print "Gouraud using vertex normals"
        Case 2
            Print "Fake half-angle specular per vertex"
        Case 3
            Print "Fake half-angle specular per face"
        Case 4
            Print "Mirror"
    End Select

    Print "frame advance"; frame_advance

    _Limit 60
    _Display

    $Checking:On
    KeyNow = UCase$(InKey$)
    If KeyNow <> "" Then

        If KeyNow = "F" Then
            T1_Filter_Selection = T1_Filter_Selection + 1
            If T1_Filter_Selection > 3 Then T1_Filter_Selection = 0
        ElseIf KeyNow = "S" Then
            Animate_Spin = Not Animate_Spin
        ElseIf KeyNow = "D" Then
            Dither_Selection = Dither_Selection + 1
            If Dither_Selection > 2 Then Dither_Selection = 0
        ElseIf KeyNow = "Z" Then
            fPitch = fPitch + 5.0
            If fPitch > 85.0 Then fPitch = 85.0
        ElseIf KeyNow = "Q" Then
            fPitch = fPitch - 5.0
            If fPitch < -85.0 Then fPitch = -85.0
        ElseIf KeyNow = "R" Then
            vCameraPsn.x = 0.0
            vCameraPsn.y = 0.0
            vCameraPsn.z = Camera_Start_Z
            fPitch = 0.0
            fYaw = 0.0
        ElseIf KeyNow = "G" Then
            Gouraud_Shading_Selection = Gouraud_Shading_Selection + 1
            If Gouraud_Shading_Selection > 4 Then Gouraud_Shading_Selection = 0
        ElseIf Asc(KeyNow) = 27 Then
            ExitCode = 1
        End If
    End If

    ' overrides
    'If Vtx_Normals_Count = 0 Then Gouraud_Shading_Selection = 0

    frametimestamp_now_ms = Timer(0.001)
    If frametimestamp_now_ms - frametimestamp_prior_ms < 0.0 Then
        ' timer rollover
        ' without over-analyzing just use the previous delta, even if it is somewhat wrong it is a better guess than 0.
        frametimestamp_prior_ms = frametimestamp_now_ms - frametimestamp_delta_ms
    Else
        frametimestamp_delta_ms = frametimestamp_now_ms - frametimestamp_prior_ms
    End If

    frame_advance = 0
    While frametimestamp_delta_ms > frametime_fullframethreshold_ms
        frame_advance = frame_advance + 1

        If _KeyDown(32) Then
            ' Spacebar
            vCameraPsn.y = vCameraPsn.y + 0.2
        End If

        If _KeyDown(118) Or _KeyDown(86) Then
            'V
            vCameraPsn.y = vCameraPsn.y - 0.2
        End If

        If _KeyDown(19712) Then
            ' Right arrow
            fYaw = fYaw - 1.2
        End If

        If _KeyDown(19200) Then
            ' Left arrow
            fYaw = fYaw + 1.2
        End If

        ' Move the player
        Matrix4_MakeRotation_Y fYaw, matCameraRot()
        Multiply_Vector3_Matrix4 vCameraHomeFwd, matCameraRot(), vMove_Player_Forward
        Vector3_Mul vMove_Player_Forward, 0.2, vMove_Player_Forward

        If _KeyDown(18432) Then
            ' Up arrow
            Vector3_Add vCameraPsn, vMove_Player_Forward, vCameraPsn
        End If

        If _KeyDown(20480) Then
            ' Down arrow
            Vector3_Delta vCameraPsn, vMove_Player_Forward, vCameraPsn
        End If

        frametimestamp_prior_ms = frametimestamp_prior_ms + frametime_fullframe_ms
        frametimestamp_delta_ms = frametimestamp_delta_ms - frametime_fullframe_ms
    Wend ' frametime

Loop Until ExitCode <> 0

For refIndex = 5 To 0 Step -1
    _FreeImage SkyBoxRef(refIndex)
Next refIndex

End
$Checking:Off

Texture1Data:
'Red_Brick', 16x16px
Data &HFF000000,&HFFdd340a,&HFFd93f0f,&HFFcfcbd2,&HFFd93f0f,&HFFdd3d1b,&HFFd93509,&HFFe04609,&HFFd93f0f,&HFFe14a1b,&HFFcf380a,&HFFd1d1d1,&HFFd93f0f,&HFFde4712,&HFFd93f0f,&HFFd93f0f
Data &HFFb33409,&HFFb33409,&HFFbe3912,&HFFdce1d7,&HFFe1501b,&HFFb33409,&HFFae2a0a,&HFFac2213,&HFFb33409,&HFFaa320f,&HFFbf3806,&HFFd6ceda,&HFFd93f0f,&HFFb73508,&HFFb33409,&HFFb43705
Data &HFFb33409,&HFFba3e19,&HFFac2903,&HFFd1d1d1,&HFFd93f0f,&HFFa93b13,&HFFb33409,&HFFa83207,&HFFb33409,&HFFb33409,&HFFb33409,&HFFd1d1d1,&HFFd12d02,&HFFc04110,&HFFb33409,&HFFbc4107
Data &HFFb89f91,&HFFb89a8a,&HFFae8e8f,&HFFd1d1d1,&HFFb29291,&HFFb3938b,&HFFb3938b,&HFFaf8f81,&HFFb3938b,&HFFbb9d8d,&HFFc3898d,&HFFd7dee4,&HFFb3938b,&HFFb3938b,&HFFaa8d8c,&HFFc8a597
Data &HFFd0320a,&HFFd93f0f,&HFFd93f0f,&HFFd93f0f,&HFFd93f0f,&HFFd73907,&HFFd93f0f,&HFFd1d1d1,&HFFde4619,&HFFd94915,&HFFe23d14,&HFFd93f0f,&HFFd93f0f,&HFFd93f0f,&HFFd93f0f,&HFFd0d0ca
Data &HFFd93f0f,&HFFb33409,&HFFc1481e,&HFFb33409,&HFFa92902,&HFFb33409,&HFFb73006,&HFFd1d1d1,&HFFd93f0f,&HFFb33409,&HFFaf2e05,&HFFb33409,&HFFac2b00,&HFFb33409,&HFFb33409,&HFFcbd9d5
Data &HFFd93f0f,&HFFa62600,&HFFb83c15,&HFFb02b00,&HFFb33409,&HFFaf2103,&HFFb33409,&HFFcfe0dd,&HFFdb4b15,&HFFaa2e02,&HFFb63b0f,&HFFad3209,&HFFb33409,&HFFb33409,&HFFb33409,&HFFd4cfd1
Data &HFFb3938b,&HFFb3938b,&HFFb3938b,&HFFaa938b,&HFFb0918e,&HFFc1998f,&HFFb3938b,&HFFd1d1d1,&HFFb3938b,&HFFb3938b,&HFFb3878b,&HFFb3938b,&HFFb3938b,&HFFb3938b,&HFFb69384,&HFFd1d1d1
Data &HFFdf3f14,&HFFd93f0f,&HFFd93f0f,&HFFd1d1d1,&HFFda3d08,&HFFd93f0f,&HFFda4515,&HFFee5519,&HFFde3f19,&HFFd93f0f,&HFFd93f0f,&HFFd1d1d1,&HFFd93f0f,&HFFd93f0f,&HFFe95407,&HFFc62f00
Data &HFFb33409,&HFFb33409,&HFFb33409,&HFFc8cfd1,&HFFde3419,&HFFb33409,&HFFb1350d,&HFFb33409,&HFFb93c0d,&HFFb64305,&HFFb92f0a,&HFFcfcacd,&HFFd93f0f,&HFFbb3a17,&HFFb33409,&HFFb33409
Data &HFFb14310,&HFFb33409,&HFFbc3b0f,&HFFd8dede,&HFFd93f0f,&HFFb33409,&HFFb0150c,&HFFb33409,&HFFb52500,&HFFb23a18,&HFFb33409,&HFFcad1d3,&HFFd93f0f,&HFFb33409,&HFFb53d0f,&HFFb33409
Data &HFFb3938b,&HFFb3938b,&HFFa58588,&HFFd1d1d1,&HFFb3938b,&HFFb3938b,&HFFb3938b,&HFFb3938b,&HFFb3938b,&HFFb3938b,&HFFad9092,&HFFd1d5d8,&HFFba8e88,&HFFb99e96,&HFFa89283,&HFFaa9784
Data &HFFca4a0f,&HFFd84118,&HFFe14a18,&HFFd93f0f,&HFFd93f0f,&HFFd93f0f,&HFFcf3d06,&HFFc7cbcf,&HFFd93b02,&HFFdd561a,&HFFd93f0f,&HFFd93f0f,&HFFd83c15,&HFFd93f0f,&HFFdd3f12,&HFFdee0e0
Data &HFFd93f0f,&HFFb33409,&HFFa93515,&HFFae3709,&HFFb33409,&HFFb33409,&HFFaa2709,&HFFc4cdcc,&HFFd93f0f,&HFFbd2d09,&HFFb33409,&HFFb33409,&HFFb33409,&HFFb43507,&HFFb73510,&HFFd1d1d1
Data &HFFd93f0f,&HFFa61800,&HFFb93a1a,&HFFb33409,&HFFb33409,&HFFb33409,&HFFb33409,&HFFd1d3d4,&HFFd93f0f,&HFFb33409,&HFFb33409,&HFFb33409,&HFFb63521,&HFFb33409,&HFFaf3c0c,&HFFd1d1d1
Data &HFFb3938b,&HFFb59181,&HFFa5857c,&HFFb59895,&HFFb3938b,&HFFb09a96,&HFFb3938b,&HFFd1d1d1,&HFFbaa085,&HFFb3938b,&HFFad8c8a,&HFFb3938b,&HFFbc9a8e,&HFFb3938b,&HFFb3938b,&HFFbec7c9

'Origin16x16', 16x16px
Data &HFFffffff,&HFFff7f27,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff
Data &HFFffdbc4,&HFFff7f27,&HFFffdbc4,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff
Data &HFFffac75,&HFFff7f27,&HFFffac75,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff
Data &HFFff7f27,&HFFff7f27,&HFFff7f27,&HFFffffff,&HFFff7f27,&HFFffffff,&HFFff7f27,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff
Data &HFFffffff,&HFFff7f27,&HFFffffff,&HFFffffff,&HFFff7f27,&HFFffffff,&HFFff7f27,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff
Data &HFFffffff,&HFFff7f27,&HFFffffff,&HFFffffff,&HFFffffff,&HFFff7f27,&HFFffdbc4,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff
Data &HFFffffff,&HFFff7f27,&HFFffffff,&HFFffffff,&HFFffffff,&HFFff7f27,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFF00a2e8,&HFFffffff,&HFF00a2e8,&HFFffffff
Data &HFFffffff,&HFFff7f27,&HFFffffff,&HFFffffff,&HFFff7f27,&HFFffdbc4,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFF00a2e8,&HFF9de1ff,&HFF00a2e8,&HFFffffff
Data &HFFffffff,&HFFff7f27,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFF00a2e8,&HFFffffff,&HFFffffff
Data &HFFffffff,&HFFff7f27,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFF00a2e8,&HFF9de1ff,&HFF00a2e8,&HFFffffff
Data &HFFffffff,&HFFff7f27,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFF00a2e8,&HFFffffff,&HFF00a2e8,&HFFffffff
Data &HFFffffff,&HFFff7f27,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff
Data &HFFffffff,&HFFff7f27,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff
Data &HFFffffff,&HFFff7f27,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFF00a2e8,&HFF64d0ff,&HFF9de1ff,&HFFffffff
Data &HFFffffff,&HFFa349a4,&HFF00a2e8,&HFF00a2e8,&HFF00a2e8,&HFF00a2e8,&HFF00a2e8,&HFF00a2e8,&HFF00a2e8,&HFF00a2e8,&HFF00a2e8,&HFF00a2e8,&HFF00a2e8,&HFF00a2e8,&HFF00a2e8,&HFF00a2e8
Data &HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFFffffff,&HFF00a2e8,&HFF64d0ff,&HFF9de1ff,&HFFffffff


' u,v texture coords use fencepost counting.
'
'   0    1    2    3    4    u
' 0 +----+----+----+----+
'   |    |    |    |    |
'   |    |    |    |    |
' 1 +----+----+----+----+
'   |    |    |    |    |
'   |    |    |    |    |
' 2 +----+----+----+----+
'   |    |    |    |    |
'   |    |    |    |    |
' 3 +----+----+----+----+
'   |    |    |    |    |
'   |    |    |    |    |
' 4 +----+----+----+----+
'
'
' v

' x0,y0,z0, x1,y1,z1, x2,y2,z2
' u0,v0, u1,v1, u2,v2
' texture_index

SKYBOX:
' FRONT Z+
Data -10,+10,+10
Data +10,+10,+10
Data -10,-10,+10
Data 0,0,128,0,0,128
Data 4

Data +10,+10,+10
Data +10,-10,+10
Data -10,-10,+10
Data 128,0,128,128,0,128
Data 4

' RIGHT X+
Data +10,+10,+10
Data +10,+10,-10
Data +10,-10,+10
Data 0,0,128,0,0,128
Data 0

Data +10,+10,-10
Data +10,-10,-10
Data +10,-10,+10
Data 128,0,128,128,0,128
Data 0

' LEFT X-
Data -10,+10,-10
Data -10,+10,+10
Data -10,-10,-10
Data 0,0,128,0,0,128
Data 1

Data -10,+10,+10
Data -10,-10,+10
Data -10,-10,-10
Data 128,0,128,128,0,128
Data 1

' BACK Z-
Data +10,+10,-10
Data -10,+10,-10
Data +10,-10,-10
Data 0,0,128,0,0,128
Data 5

Data -10,+10,-10
Data -10,-10,-10
Data +10,-10,-10
Data 128,0,128,128,0,128
Data 5

' TOP Y+
Data -10,+10,-10
Data +10,+10,-10
Data -10,+10,+10
Data 0,0,128,0,0,128
Data 2

Data +10,+10,-10
Data +10,+10,+10
Data -10,+10,+10
Data 128,0,128,128,0,128
Data 2

' BOTTOM Y-
Data -10,-10,+10
Data 10,-10,+10
Data -10,-10,-10
Data 0,0,128,0,0,128
Data 3

Data +10,-10,+10
Data +10,-10,-10
Data -10,-10,-10
Data 128,0,128,128,0,128
Data 3

$Checking:On
Sub PrescanMesh (thefile As String, requiredTriangles As Long, totalVertex As Long, totalTextureCoords As Long, totalNormals As Long, totalMaterialLibrary As Long, materialFile As String)
    ' Primary purpose is to determine the required amount of triangles.
    ' This is not straightforward as any size n-gons are allowed, although typically faces with 3 or 4 vertexes is encountered.
    Dim totalFaces As Long
    Dim lineCount As Long
    Dim lineLength As Integer
    Dim lineCursor As Integer
    Dim parameterIndex As Integer
    Dim substringStart As Integer
    Dim substringLength As Integer

    Dim rawString As String

    requiredTriangles = 0
    totalVertex = 0
    totalTextureCoords = 0
    totalNormals = 0
    totalMaterialLibrary = 0
    materialFile = "default.mtl"

    totalFaces = 0
    lineCount = 0
    lineLength = 0

    If _FileExists(thefile) = 0 Then
        Print "The file was not found"
        Exit Sub
    End If
    Open thefile For Binary As #2
    If LOF(2) = 0 Then
        Print "The file is empty"
        Close #2
        Exit Sub
    End If

    Do Until EOF(2)
        Line Input #2, rawString
        lineCount = lineCount + 1

        If Left$(rawString, 2) = "v " Then
            totalVertex = totalVertex + 1

        ElseIf Left$(rawString, 3) = "vt " Then
            totalTextureCoords = totalTextureCoords + 1

        ElseIf Left$(rawString, 3) = "vn " Then
            totalNormals = totalNormals + 1


        ElseIf Left$(rawString, 2) = "f " Then
            totalFaces = totalFaces + 1

            lineCursor = 3
            parameterIndex = 0

            lineLength = Len(rawString)
            While lineCursor < lineLength
                substringStart = lineCursor
                substringLength = 0

                ' eat spaces
                While Asc(rawString, lineCursor) <= 32
                    lineCursor = lineCursor + 1
                    If lineCursor > lineLength Then GoTo BAIL_FACE_PRESCAN
                Wend

                ' count number digits
                While Asc(rawString, lineCursor) > 32
                    substringLength = substringLength + 1
                    lineCursor = lineCursor + 1
                    If lineCursor > lineLength Then Exit While
                Wend

                If substringLength > 0 Then
                    parameterIndex = parameterIndex + 1
                End If
            Wend

            BAIL_FACE_PRESCAN:
            If parameterIndex = 3 Then
                requiredTriangles = requiredTriangles + 1
            Else
                requiredTriangles = requiredTriangles + 2
            End If

        ElseIf Left$(rawString, 7) = "mtllib " Then
            totalMaterialLibrary = totalMaterialLibrary + 1
            ' return the most recently encountered material file. there is usually only one.
            materialFile = _Trim$(Mid$(rawString, 8))
        End If
    Loop

    Close #2

    Print lineCount, "lineCount"
    Print requiredTriangles, "Required Triangles"
    Print totalVertex, "Vertexes"
    Print totalTextureCoords, "Texture Coordinates"
    Print totalNormals, "Vertex Normals"
    Print totalFaces, "Faces"
    Print totalMaterialLibrary, "Material Libraries"
    Print materialFile

    'Dim temp$
    'Input "waiting..."; temp$
End Sub

Sub LoadMesh (thefile As String, tris() As triangle, indexTri As Long, leVertexList As Long, leVertexTexelList As Long, mats() As newmtl_type, vn() As vec3d)
    Dim ParameterStorage(10, 2) As Double

    Dim totalVertex As Long
    Dim totalFaces As Long
    Dim lineCount As Long
    Dim lineLength As Integer
    Dim lineCursor As Integer
    Dim parameterIndex As Integer
    Dim paramStringStart As Integer
    Dim paramStringLength As Integer
    Dim useMaterialNumber As Long
    Dim totalVertexNormals As Long
    Dim totalVertexTexels As Long
    Dim VertexList(leVertexList) As vec3d ' this gets tossed after loading mesh()
    Dim TexelCoord(leVertexTexelList) As vec3d ' this gets tossed after loading mesh()

    totalVertex = 0
    totalVertexNormals = 0 ' refers to vn() array

    lineCount = 0
    lineLength = 0

    Dim rawString As String
    Dim parameter As String

    If _FileExists(thefile) = 0 Then
        Print "The file was not found"
        Exit Sub
    End If
    Open thefile For Binary As #2
    If LOF(2) = 0 Then
        Print "The file is empty"
        Close #2
        Exit Sub
    End If
    Do Until EOF(2)
        Line Input #2, rawString
        lineCount = lineCount + 1

        lineLength = Len(rawString)

        If Left$(rawString, 2) = "v " Then
            totalVertex = totalVertex + 1
            'Print "Vertex #"; totalVertex;

            lineCursor = 3
            parameterIndex = 0

            While lineCursor < lineLength
                paramStringLength = 0

                ' eat spaces
                While Asc(rawString, lineCursor) <= 32
                    lineCursor = lineCursor + 1
                    If lineCursor > lineLength Then GoTo BAIL_VERTEX
                Wend

                ' count chars up to next space char, stopping if end of string reached
                paramStringStart = lineCursor
                While Asc(rawString, lineCursor) > 32
                    'Print Mid$(rawString, lineCursor, 1);
                    paramStringLength = paramStringLength + 1
                    lineCursor = lineCursor + 1
                    If lineCursor > lineLength Then Exit While
                Wend

                If paramStringLength > 0 Then
                    parameterIndex = parameterIndex + 1
                    parameter$ = Mid$(rawString, paramStringStart, paramStringLength)
                    ParameterStorage(parameterIndex, 0) = Val(parameter$)
                    'Print parameterIndex, paramStringStart, paramStringLength, "["; parameter$; "]"
                End If
            Wend

            BAIL_VERTEX:
            'Print parameterIndex; " EOL"

            If parameterIndex >= 3 Then
                VertexList(totalVertex).x = ParameterStorage(1, 0)
                VertexList(totalVertex).y = ParameterStorage(2, 0)
                VertexList(totalVertex).z = ParameterStorage(3, 0)
                'Print "v "; ParameterStorage(1); ParameterStorage(2); ParameterStorage(3)
            End If

        ElseIf Left$(rawString, 3) = "vn " Then
            totalVertexNormals = totalVertexNormals + 1

            lineCursor = 4
            parameterIndex = 0

            While lineCursor < lineLength
                paramStringLength = 0

                ' eat spaces
                While Asc(rawString, lineCursor) <= 32
                    lineCursor = lineCursor + 1
                    If lineCursor > lineLength Then GoTo BAIL_VERTEXNORMS
                Wend

                ' count chars up to next space char, stopping if end of string reached
                paramStringStart = lineCursor
                While Asc(rawString, lineCursor) > 32
                    'Print Mid$(rawString, lineCursor, 1);
                    paramStringLength = paramStringLength + 1
                    lineCursor = lineCursor + 1
                    If lineCursor > lineLength Then Exit While
                Wend

                If paramStringLength > 0 Then
                    parameterIndex = parameterIndex + 1
                    parameter$ = Mid$(rawString, paramStringStart, paramStringLength)
                    ParameterStorage(parameterIndex, 0) = Val(parameter$)
                    'Print parameterIndex, paramStringStart, paramStringLength, "["; parameter$; "]"
                End If
            Wend

            BAIL_VERTEXNORMS:
            'Print parameterIndex; " EOL"

            If parameterIndex >= 3 Then
                vn(totalVertexNormals).x = ParameterStorage(1, 0)
                vn(totalVertexNormals).y = ParameterStorage(2, 0)
                vn(totalVertexNormals).z = ParameterStorage(3, 0)
                Vector3_Normalize vn(totalVertexNormals)
                'Print "vn "; ParameterStorage(1); ParameterStorage(2); ParameterStorage(3)
            End If


        ElseIf Left$(rawString, 3) = "vt " Then
            totalVertexTexels = totalVertexTexels + 1

            lineCursor = 4
            parameterIndex = 0

            While lineCursor < lineLength
                paramStringLength = 0

                ' eat spaces
                While Asc(rawString, lineCursor) <= 32
                    lineCursor = lineCursor + 1
                    If lineCursor > lineLength Then GoTo BAIL_VERTEXTEXELS
                Wend

                ' count chars up to next space char, stopping if end of string reached
                paramStringStart = lineCursor
                While Asc(rawString, lineCursor) > 32
                    'Print Mid$(rawString, lineCursor, 1);
                    paramStringLength = paramStringLength + 1
                    lineCursor = lineCursor + 1
                    If lineCursor > lineLength Then Exit While
                Wend

                If paramStringLength > 0 Then
                    parameterIndex = parameterIndex + 1
                    parameter$ = Mid$(rawString, paramStringStart, paramStringLength)
                    ParameterStorage(parameterIndex, 0) = Val(parameter$)
                    'Print parameterIndex, paramStringStart, paramStringLength, "["; parameter$; "]"
                End If
            Wend

            BAIL_VERTEXTEXELS:
            'Print parameterIndex; " EOL"

            If parameterIndex >= 2 Then
                TexelCoord(totalVertexTexels).x = ParameterStorage(1, 0)
                TexelCoord(totalVertexTexels).y = ParameterStorage(2, 0)
                'Print "vt "; ParameterStorage(1, 0); ParameterStorage(2, 0);
            End If

        End If
    Loop

    lineCount = 0
    lineLength = 0
    totalFaces = 0

    Dim mostRecentVertex As Long
    Dim mostRecentVertexNormal As Long
    Dim mostRecentVertexTexel As Long
    mostRecentVertex = 0
    mostRecentVertexNormal = 0
    mostRecentVertexTexel = 0

    Dim tri0 As Long
    Dim tri1 As Long
    Dim tri2 As Long
    Dim tri3 As Long
    Dim tex0 As Long
    Dim tex1 As Long
    Dim tex2 As Long
    Dim tex3 As Long

    useMaterialNumber = 0

    Dim ssc As Integer
    ssc = 0

    Dim paramSubindex As Integer
    paramSubindex = 0

    Seek #2, 1
    Do Until EOF(2)
        Line Input #2, rawString
        lineCount = lineCount + 1
        lineLength = Len(rawString)

        If Left$(rawString, 2) = "v " Then
            ' vertex command
            ' Increase vertex counter due to the relative vertex feature
            ' If a vertex in a face command is -1, it refers to the most recent vertex.
            ' So then -2 means the second most recent vertex, etc.
            mostRecentVertex = mostRecentVertex + 1

        ElseIf Left$(rawString, 3) = "vn " Then
            ' vertex normal command (directional shading)
            mostRecentVertexNormal = mostRecentVertexNormal + 1

        ElseIf Left$(rawString, 3) = "vt " Then
            ' vertex texture coordinate(s)
            mostRecentVertexTexel = mostRecentVertexTexel + 1

        ElseIf Left$(rawString, 2) = "f " Then
            ' face command
            ' f v1/vt1/vn1 v2/vt2/vn2 v3/vt3/vn3 ...
            ' f vt//vn1 v2//vn2 v3//vn3 ...
            totalFaces = totalFaces + 1
            'Print "Face #"; totalFaces;

            lineCursor = 3
            parameterIndex = 0
            paramSubindex = 0

            While lineCursor < lineLength
                paramStringLength = 0

                ' eat spaces
                ' space is used as a separator between parameters
                While Asc(rawString, lineCursor) <= 32
                    lineCursor = lineCursor + 1
                    If lineCursor > lineLength Then GoTo BAIL_FACES
                    paramSubindex = 0
                Wend

                ' count chars up to next space or slash char, stopping if end of string reached
                paramStringStart = lineCursor
                ssc = Asc(rawString, lineCursor)
                While (ssc > 32)
                    lineCursor = lineCursor + 1
                    If ssc = 47 Then Exit While
                    paramStringLength = paramStringLength + 1
                    If lineCursor > lineLength Then Exit While
                    ssc = Asc(rawString, lineCursor)
                Wend

                If paramStringLength > 0 Then
                    If paramSubindex = 0 Then parameterIndex = parameterIndex + 1
                    parameter$ = Mid$(rawString, paramStringStart, paramStringLength)
                    ParameterStorage(parameterIndex, paramSubindex) = Val(parameter$)
                    'Print "("; parameterIndex; ","; paramSubindex; ")", paramStringStart, paramStringLength, "["; parameter$; "]"
                End If

                If ssc = 47 Then
                    ' / character means next sub parameter
                    paramSubindex = paramSubindex + 1
                    ParameterStorage(parameterIndex, paramSubindex) = 0
                End If

            Wend

            BAIL_FACES:
            'Print parameterIndex; " EOL"
            If parameterIndex >= 3 Then
                tri0 = FindVertexNumAbsOrRel(ParameterStorage(1, 0), mostRecentVertex)
                tri1 = FindVertexNumAbsOrRel(ParameterStorage(2, 0), mostRecentVertex)
                tri2 = FindVertexNumAbsOrRel(ParameterStorage(3, 0), mostRecentVertex)

                indexTri = indexTri + 1
                tris(indexTri).x0 = VertexList(tri0).x
                tris(indexTri).y0 = VertexList(tri0).y
                tris(indexTri).z0 = VertexList(tri0).z
                tris(indexTri).x1 = VertexList(tri1).x
                tris(indexTri).y1 = VertexList(tri1).y
                tris(indexTri).z1 = VertexList(tri1).z
                tris(indexTri).x2 = VertexList(tri2).x
                tris(indexTri).y2 = VertexList(tri2).y
                tris(indexTri).z2 = VertexList(tri2).z
                tris(indexTri).options = 0
                tris(indexTri).material = useMaterialNumber

                If paramSubindex >= 1 Then
                    If ParameterStorage(1, 1) = 0 Then
                        ' no texture 1
                        tris(indexTri).options = tris(indexTri).options Or T1_option_no_T1
                    Else
                        tex0 = FindVertexNumAbsOrRel(ParameterStorage(1, 1), mostRecentVertexTexel)
                        tex1 = FindVertexNumAbsOrRel(ParameterStorage(2, 1), mostRecentVertexTexel)
                        tex2 = FindVertexNumAbsOrRel(ParameterStorage(3, 1), mostRecentVertexTexel)

                        tris(indexTri).u0 = TexelCoord(tex0).x
                        tris(indexTri).v0 = TexelCoord(tex0).y

                        tris(indexTri).u1 = TexelCoord(tex1).x
                        tris(indexTri).v1 = TexelCoord(tex1).y

                        tris(indexTri).u2 = TexelCoord(tex2).x
                        tris(indexTri).v2 = TexelCoord(tex2).y
                    End If
                End If

                If paramSubindex >= 2 Then
                    tris(indexTri).vni0 = FindVertexNumAbsOrRel(ParameterStorage(1, 2), mostRecentVertexNormal)
                    tris(indexTri).vni1 = FindVertexNumAbsOrRel(ParameterStorage(2, 2), mostRecentVertexNormal)
                    tris(indexTri).vni2 = FindVertexNumAbsOrRel(ParameterStorage(3, 2), mostRecentVertexNormal)
                Else
                    tris(indexTri).vni0 = 0
                    tris(indexTri).vni1 = 0
                    tris(indexTri).vni2 = 0
                End If
            End If

            If parameterIndex = 4 Then
                tri3 = FindVertexNumAbsOrRel(ParameterStorage(4, 0), mostRecentVertex)
                tex3 = FindVertexNumAbsOrRel(ParameterStorage(4, 1), mostRecentVertexTexel)

                indexTri = indexTri + 1
                tris(indexTri).x0 = VertexList(tri0).x
                tris(indexTri).y0 = VertexList(tri0).y
                tris(indexTri).z0 = VertexList(tri0).z
                tris(indexTri).x1 = VertexList(tri2).x
                tris(indexTri).y1 = VertexList(tri2).y
                tris(indexTri).z1 = VertexList(tri2).z
                tris(indexTri).x2 = VertexList(tri3).x
                tris(indexTri).y2 = VertexList(tri3).y
                tris(indexTri).z2 = VertexList(tri3).z
                tris(indexTri).options = 0
                tris(indexTri).material = useMaterialNumber

                If paramSubindex >= 1 Then
                    If (tris(indexTri).options And T1_option_no_T1) = 0 Then
                        tris(indexTri).u0 = TexelCoord(tex0).x
                        tris(indexTri).v0 = TexelCoord(tex0).y

                        tris(indexTri).u1 = TexelCoord(tex2).x
                        tris(indexTri).v1 = TexelCoord(tex2).y

                        tris(indexTri).u2 = TexelCoord(tex3).x
                        tris(indexTri).v2 = TexelCoord(tex3).y
                    End If
                End If

                If paramSubindex >= 2 Then
                    tris(indexTri).vni0 = FindVertexNumAbsOrRel(ParameterStorage(1, 2), mostRecentVertexNormal)
                    tris(indexTri).vni1 = FindVertexNumAbsOrRel(ParameterStorage(3, 2), mostRecentVertexNormal)
                    tris(indexTri).vni2 = FindVertexNumAbsOrRel(ParameterStorage(4, 2), mostRecentVertexNormal)
                Else
                    tris(indexTri).vni0 = 0
                    tris(indexTri).vni1 = 0
                    tris(indexTri).vni2 = 0
                End If
            End If

            If (parameterIndex <> 3) And (parameterIndex <> 4) Then
                Print "Line "; lineCount; " what kind of face is this? ***************************************"
                Sleep 3
            End If

        ElseIf Left$(rawString, 7) = "usemtl " Then
            ' use material command

            lineCursor = 8
            parameterIndex = 0

            While lineCursor < lineLength
                paramStringLength = 0

                ' eat spaces
                While Asc(rawString, lineCursor) <= 32
                    lineCursor = lineCursor + 1
                    If lineCursor > lineLength Then GoTo BAIL_SETMTL
                Wend

                ' count chars up to next space char, stopping if end of string reached
                paramStringStart = lineCursor
                While Asc(rawString, lineCursor) > 32
                    paramStringLength = paramStringLength + 1
                    lineCursor = lineCursor + 1
                    If lineCursor > lineLength Then Exit While
                Wend

                If paramStringLength > 0 Then
                    parameterIndex = parameterIndex + 1
                    parameter$ = Mid$(rawString, paramStringStart, paramStringLength)
                    useMaterialNumber = Find_Material_id_from_name(mats(), parameter$)
                    Print "use material ["; parameter$; "]", useMaterialNumber
                End If
            Wend

            BAIL_SETMTL:
        End If
    Loop
    Close #2
    'Dim temp$
    'Input "waiting..."; temp$
    'Sleep 3
End Sub

Function FindVertexNumAbsOrRel (param As Double, recentvert As Long)
    If param < 0 Then
        ' relative
        ' -1 means the most recent vertex encountered
        FindVertexNumAbsOrRel = param + recentvert + 1
    Else
        FindVertexNumAbsOrRel = param
    End If
End Function

Sub PrescanMaterialFile (thefile As String, totalMaterials As Long)
    Dim lineCount As Long
    Dim rawString As String
    Dim trimString As String

    totalMaterials = 0
    lineCount = 0

    If _FileExists(thefile) = 0 Then
        Print "The file was not found"
        Exit Sub
    End If
    Open thefile For Binary As #2
    If LOF(2) = 0 Then
        Print "The file is empty"
        Close #2
        Exit Sub
    End If

    Do Until EOF(2)
        Line Input #2, rawString
        lineCount = lineCount + 1
        trimString = _Trim$(rawString)

        If Left$(trimString, 7) = "newmtl " Then
            totalMaterials = totalMaterials + 1
        End If
    Loop
    Close #2

    Print totalMaterials, "Total Materials"
    'Dim temp$
    'Input "waiting..."; temp$

End Sub

Sub LoadMaterialFile (theFile As String, mats() As newmtl_type, totalMaterials As Long)
    Dim ParameterStorage(10) As Double

    Dim lineCount As Long
    Dim lineLength As Integer
    Dim lineCursor As Integer
    Dim parameterIndex As Integer
    Dim substringStart As Integer
    Dim substringLength As Integer

    Dim rawString As String
    Dim trimString As String
    Dim parameter As String

    totalMaterials = 0
    lineCount = 0

    If _FileExists(theFile) = 0 Then
        Print "The file was not found"
        Exit Sub
    End If
    Open theFile For Binary As #2
    If LOF(2) = 0 Then
        Print "The file is empty"
        Close #2
        Exit Sub
    End If

    Do Until EOF(2)
        lineCount = lineCount + 1
        Line Input #2, rawString
        lineLength = Len(rawString)
        If lineLength < 1 Then _Continue

        ' eat spaces
        ' unfortunately there is a tendency to have tabs or spaces in front of parameters to indent them
        lineCursor = 1
        While Asc(rawString, lineCursor) <= 32
            lineCursor = lineCursor + 1
            If lineCursor > lineLength Then GoTo BAIL_INDENT_MATERIAL
        Wend
        trimString = Mid$(rawString, lineCursor)

        lineCursor = 1 'now looking at trimString
        lineLength = Len(trimString)
        If Left$(trimString, 7) = "newmtl " Then
            totalMaterials = totalMaterials + 1
            mats(totalMaterials).textName = _Trim$(Mid$(trimString, 8))
            ' set any defaults that are non-zero
            mats(totalMaterials).diaphaneity = 1.0

        ElseIf Left$(trimString, 3) = "Kd " Then
            ' Diffuse color is the most dominant color
            lineCursor = 3
            parameterIndex = 0

            While lineCursor < lineLength
                substringLength = 0

                ' eat spaces
                While Asc(trimString, lineCursor) <= 32
                    lineCursor = lineCursor + 1
                    If lineCursor > lineLength Then GoTo BAIL_KD
                Wend

                ' count chars up to next space char, stopping if end of string reached
                substringStart = lineCursor
                While Asc(trimString, lineCursor) > 32
                    'Print Mid$(trimString, lineCursor, 1);
                    substringLength = substringLength + 1
                    lineCursor = lineCursor + 1
                    If lineCursor > lineLength Then Exit While
                Wend

                If substringLength > 0 Then
                    parameterIndex = parameterIndex + 1
                    parameter$ = Mid$(trimString, substringStart, substringLength)
                    ParameterStorage(parameterIndex) = Val(parameter$)
                    'Print parameterIndex, substringStart, substringLength, "["; parameter$; "]"
                End If
            Wend

            BAIL_KD:
            If parameterIndex = 3 Then
                mats(totalMaterials).Kd_r = ParameterStorage(1)
                mats(totalMaterials).Kd_g = ParameterStorage(2)
                mats(totalMaterials).Kd_b = ParameterStorage(3)
            End If
            Print totalMaterials, "["; Mid$(trimString, 4); "]"

        ElseIf Left$(trimString, 3) = "Ks " Then
            ' Specular color
            lineCursor = 3
            parameterIndex = 0

            While lineCursor < lineLength
                substringLength = 0

                ' eat spaces
                While Asc(trimString, lineCursor) <= 32
                    lineCursor = lineCursor + 1
                    If lineCursor > lineLength Then GoTo BAIL_KS
                Wend

                ' count chars up to next space char, stopping if end of string reached
                substringStart = lineCursor
                While Asc(trimString, lineCursor) > 32
                    'Print Mid$(trimString, lineCursor, 1);
                    substringLength = substringLength + 1
                    lineCursor = lineCursor + 1
                    If lineCursor > lineLength Then Exit While
                Wend

                If substringLength > 0 Then
                    parameterIndex = parameterIndex + 1
                    parameter$ = Mid$(trimString, substringStart, substringLength)
                    ParameterStorage(parameterIndex) = Val(parameter$)
                    'Print parameterIndex, substringStart, substringLength, "["; parameter$; "]"
                End If
            Wend

            BAIL_KS:
            If parameterIndex = 3 Then
                mats(totalMaterials).Ks_r = ParameterStorage(1)
                mats(totalMaterials).Ks_g = ParameterStorage(2)
                mats(totalMaterials).Ks_b = ParameterStorage(3)
            End If
            'Print totalMaterials, "["; Mid$(trimString, 4); "]"

        ElseIf Left$(trimString, 2) = "d " Then
            ' diaphaneity
            mats(totalMaterials).diaphaneity = Val(Mid$(trimString, 3))

        End If
        BAIL_INDENT_MATERIAL:
    Loop
    Close #2

    Print totalMaterials, "Total Materials"
    Dim i As Long
    For i = 1 To totalMaterials
        Print i, mats(i).textName; mats(i).Kd_r; mats(i).Kd_g; mats(i).Kd_b; mats(i).diaphaneity
    Next i
    'Dim temp$
    'Input "waiting..."; temp$

End Sub

Function Find_Material_id_from_name (mats() As newmtl_type, in_name As String)
    Dim i As Long
    Find_Material_id_from_name = 0

    'Print "find ["; in_name; "]"

    For i = LBound(mats) To UBound(mats)
        'Print "{"; mats(i).textName; "} ";
        If in_name = mats(i).textName Then
            Find_Material_id_from_name = i
            'Print i; "match", in_name, mats(i).textName
            Exit For
        End If
    Next i
End Function

$Checking:Off
Sub NearClip (A As vec3d, B As vec3d, C As vec3d, D As vec3d, TA As vertex_attribute7, TB As vertex_attribute7, TC As vertex_attribute7, TD As vertex_attribute7, result As Integer)
    ' This function clips a triangle to Frustum_Near
    ' Winding order is preserved.
    ' result:
    ' 0 = do not draw
    ' 1 = only draw ABCA
    ' 2 = draw both ABCA and ACDA

    Static d_A_near_z As Single
    Static d_B_near_z As Single
    Static d_C_near_z As Single
    Static clip_score As _Unsigned Integer
    Static ratio1 As Single
    Static ratio2 As Single

    d_A_near_z = A.z - Frustum_Near
    d_B_near_z = B.z - Frustum_Near
    d_C_near_z = C.z - Frustum_Near

    clip_score = 0
    If d_A_near_z < 0.0 Then clip_score = clip_score Or 1
    If d_B_near_z < 0.0 Then clip_score = clip_score Or 2
    If d_C_near_z < 0.0 Then clip_score = clip_score Or 4

    'Print clip_score;

    Select Case clip_score
        Case &B000
            'Print "no clip"
            result = 1


        Case &B001
            'Print "A is out"
            result = 2

            ' C to new D (using C to A)
            ratio1 = d_C_near_z / (C.z - A.z)
            D.x = (A.x - C.x) * ratio1 + C.x
            D.y = (A.y - C.y) * ratio1 + C.y
            D.z = Frustum_Near
            TD.u = (TA.u - TC.u) * ratio1 + TC.u
            TD.v = (TA.v - TC.v) * ratio1 + TC.v
            TD.r = (TA.r - TC.r) * ratio1 + TC.r
            TD.g = (TA.g - TC.g) * ratio1 + TC.g
            TD.b = (TA.b - TC.b) * ratio1 + TC.b
            TD.s = (TA.s - TC.s) * ratio1 + TC.s
            TD.t = (TA.t - TC.t) * ratio1 + TC.t

            ' new A to B, going backward from B
            ratio2 = d_B_near_z / (B.z - A.z)
            A.x = (A.x - B.x) * ratio2 + B.x
            A.y = (A.y - B.y) * ratio2 + B.y
            A.z = Frustum_Near
            TA.u = (TA.u - TB.u) * ratio2 + TB.u
            TA.v = (TA.v - TB.v) * ratio2 + TB.v
            TA.r = (TA.r - TB.r) * ratio2 + TB.r
            TA.g = (TA.g - TB.g) * ratio2 + TB.g
            TA.b = (TA.b - TB.b) * ratio2 + TB.b
            TA.s = (TA.s - TB.s) * ratio2 + TB.s
            TA.t = (TA.t - TB.t) * ratio2 + TB.t


        Case &B010
            'Print "B is out"
            result = 2

            ' the oddball case
            D = C
            TD = TC

            ' old B to new C, going backward from C to B
            ratio1 = d_C_near_z / (C.z - B.z)
            C.x = (B.x - C.x) * ratio1 + C.x
            C.y = (B.y - C.y) * ratio1 + C.y
            C.z = Frustum_Near
            TC.u = (TB.u - TC.u) * ratio1 + TC.u
            TC.v = (TB.v - TC.v) * ratio1 + TC.v
            TC.r = (TB.r - TC.r) * ratio1 + TC.r
            TC.g = (TB.g - TC.g) * ratio1 + TC.g
            TC.b = (TB.b - TC.b) * ratio1 + TC.b
            TC.s = (TB.s - TC.s) * ratio1 + TC.s
            TC.t = (TB.t - TC.t) * ratio1 + TC.t

            ' A to new B, going forward from A
            ratio2 = d_A_near_z / (A.z - B.z)
            B.x = (B.x - A.x) * ratio2 + A.x
            B.y = (B.y - A.y) * ratio2 + A.y
            B.z = Frustum_Near
            TB.u = (TB.u - TA.u) * ratio2 + TA.u
            TB.v = (TB.v - TA.v) * ratio2 + TA.v
            TB.r = (TB.r - TA.r) * ratio2 + TA.r
            TB.g = (TB.g - TA.g) * ratio2 + TA.g
            TB.b = (TB.b - TA.b) * ratio2 + TA.b
            TB.s = (TB.s - TA.s) * ratio2 + TA.s
            TB.t = (TB.t - TA.t) * ratio2 + TA.t


        Case &B011
            'Print "C is in"
            result = 1

            ' new B to C
            ratio1 = d_C_near_z / (C.z - B.z)
            B.x = (B.x - C.x) * ratio1 + C.x
            B.y = (B.y - C.y) * ratio1 + C.y
            B.z = Frustum_Near
            TB.u = (TB.u - TC.u) * ratio1 + TC.u
            TB.v = (TB.v - TC.v) * ratio1 + TC.v
            TB.r = (TB.r - TC.r) * ratio1 + TC.r
            TB.g = (TB.g - TC.g) * ratio1 + TC.g
            TB.b = (TB.b - TC.b) * ratio1 + TC.b
            TB.s = (TB.s - TC.s) * ratio1 + TC.s
            TB.t = (TB.t - TC.t) * ratio1 + TC.t

            ' C to new A
            ratio2 = d_C_near_z / (C.z - A.z)
            A.x = (A.x - C.x) * ratio2 + C.x
            A.y = (A.y - C.y) * ratio2 + C.y
            A.z = Frustum_Near
            TA.u = (TA.u - TC.u) * ratio2 + TC.u
            TA.v = (TA.v - TC.v) * ratio2 + TC.v
            TA.r = (TA.r - TC.r) * ratio2 + TC.r
            TA.g = (TA.g - TC.g) * ratio2 + TC.g
            TA.b = (TA.b - TC.b) * ratio2 + TC.b
            TA.s = (TA.s - TC.s) * ratio2 + TC.s
            TA.t = (TA.t - TC.t) * ratio2 + TC.t


        Case &B100
            'Print "C is out"
            result = 2

            ' new D to A
            ratio1 = d_A_near_z / (A.z - C.z)
            D.x = (C.x - A.x) * ratio1 + A.x
            D.y = (C.y - A.y) * ratio1 + A.y
            D.z = Frustum_Near
            TD.u = (TC.u - TA.u) * ratio1 + TA.u
            TD.v = (TC.v - TA.v) * ratio1 + TA.v
            TD.r = (TC.r - TA.r) * ratio1 + TA.r
            TD.g = (TC.g - TA.g) * ratio1 + TA.g
            TD.b = (TC.b - TA.b) * ratio1 + TA.b
            TD.s = (TC.s - TA.s) * ratio1 + TA.s
            TD.t = (TC.t - TA.t) * ratio1 + TA.t

            ' B to new C
            ratio2 = d_B_near_z / (B.z - C.z)
            C.x = (C.x - B.x) * ratio2 + B.x
            C.y = (C.y - B.y) * ratio2 + B.y
            C.z = Frustum_Near
            TC.u = (TC.u - TB.u) * ratio2 + TB.u
            TC.v = (TC.v - TB.v) * ratio2 + TB.v
            TC.r = (TC.r - TB.r) * ratio2 + TB.r
            TC.g = (TC.g - TB.g) * ratio2 + TB.g
            TC.b = (TC.b - TB.b) * ratio2 + TB.b
            TC.s = (TC.s - TB.s) * ratio2 + TB.s
            TC.t = (TC.t - TB.t) * ratio2 + TB.t


        Case &B101
            'Print "B is in"
            result = 1

            ' new A to B
            ratio1 = d_B_near_z / (B.z - A.z)
            A.x = (A.x - B.x) * ratio1 + B.x
            A.y = (A.y - B.y) * ratio1 + B.y
            A.z = Frustum_Near
            TA.u = (TA.u - TB.u) * ratio1 + TB.u
            TA.v = (TA.v - TB.v) * ratio1 + TB.v
            TA.r = (TA.r - TB.r) * ratio1 + TB.r
            TA.g = (TA.g - TB.g) * ratio1 + TB.g
            TA.b = (TA.b - TB.b) * ratio1 + TB.b
            TA.s = (TA.s - TB.s) * ratio1 + TB.s
            TA.t = (TA.t - TB.t) * ratio1 + TB.t

            ' B to new C
            ratio2 = d_B_near_z / (B.z - C.z)
            C.x = (C.x - B.x) * ratio2 + B.x
            C.y = (C.y - B.y) * ratio2 + B.y
            C.z = Frustum_Near
            TC.u = (TC.u - TB.u) * ratio2 + TB.u
            TC.v = (TC.v - TB.v) * ratio2 + TB.v
            TC.r = (TC.r - TB.r) * ratio2 + TB.r
            TC.g = (TC.g - TB.g) * ratio2 + TB.g
            TC.b = (TC.b - TB.b) * ratio2 + TB.b
            TC.s = (TC.s - TB.s) * ratio2 + TB.s
            TC.t = (TC.t - TB.t) * ratio2 + TB.t


        Case &B110
            'Print "A is in"
            result = 1

            ' A to new B
            ratio1 = d_A_near_z / (A.z - B.z)
            B.x = (B.x - A.x) * ratio1 + A.x
            B.y = (B.y - A.y) * ratio1 + A.y
            B.z = Frustum_Near
            TB.u = (TB.u - TA.u) * ratio1 + TA.u
            TB.v = (TB.v - TA.v) * ratio1 + TA.v
            TB.r = (TB.r - TA.r) * ratio1 + TA.r
            TB.g = (TB.g - TA.g) * ratio1 + TA.g
            TB.b = (TB.b - TA.b) * ratio1 + TA.b
            TB.s = (TB.s - TA.s) * ratio1 + TA.s
            TB.t = (TB.t - TA.t) * ratio1 + TA.t

            ' new C to A
            ratio2 = d_A_near_z / (A.z - C.z)
            C.x = (C.x - A.x) * ratio2 + A.x
            C.y = (C.y - A.y) * ratio2 + A.y
            C.z = Frustum_Near
            TC.u = (TC.u - TA.u) * ratio2 + TA.u
            TC.v = (TC.v - TA.v) * ratio2 + TA.v
            TC.r = (TC.r - TA.r) * ratio2 + TA.r
            TC.g = (TC.g - TA.g) * ratio2 + TA.g
            TC.b = (TC.b - TA.b) * ratio2 + TA.b
            TC.s = (TC.s - TA.s) * ratio2 + TA.s
            TC.t = (TC.t - TA.t) * ratio2 + TA.t


        Case &B111
            'Print "discard"
            result = 0

    End Select

End Sub


' Multiply a 3D vector into a 4x4 matrix and output another 3D vector
' Important!: matrix o must be a different variable from matrix i. if i and o are the same variable it will malfunction.
' To understand the optimization here. Mathematically you can only multiply matrices of the same dimension. 4 here.
' But I'm only interested in x, y, and z; so don't bother calculating "w" because it is always 1.
' Avoiding 7 unnecessary extra multiplications
Sub Multiply_Vector3_Matrix4 (i As vec3d, m( 3 , 3) As Single, o As vec3d)
    o.x = i.x * m(0, 0) + i.y * m(1, 0) + i.z * m(2, 0) + m(3, 0)
    o.y = i.x * m(0, 1) + i.y * m(1, 1) + i.z * m(2, 1) + m(3, 1)
    o.z = i.x * m(0, 2) + i.y * m(1, 2) + i.z * m(2, 2) + m(3, 2)
End Sub

' If you ever need it, then go ahead and uncomment and pass in vec4d, but I doubt you ever would need it.
'SUB MultiplyMatrixVector4 (i AS vec4d, o AS vec4d, m( 3 , 3) AS SINGLE)
'    o.x = i.x * m(0, 0) + i.y * m(1, 0) + i.z * m(2, 0) + m(3, 0) * i.w
'    o.y = i.x * m(0, 1) + i.y * m(1, 1) + i.z * m(2, 1) + m(3, 1) * i.w
'    o.z = i.x * m(0, 2) + i.y * m(1, 2) + i.z * m(2, 2) + m(3, 2) * i.w
'    o.w = i.x * m(0, 3) + i.y * m(1, 3) + i.z * m(2, 3) + m(3, 3) * i.w
'END SUB

Sub Vector3_Add (left As vec3d, right As vec3d, o As vec3d) Static
    o.x = left.x + right.x
    o.y = left.y + right.y
    o.z = left.z + right.z
End Sub

Sub Vector3_Delta (left As vec3d, right As vec3d, o As vec3d) Static
    o.x = left.x - right.x
    o.y = left.y - right.y
    o.z = left.z - right.z
End Sub

Sub Vector3_Normalize (io As vec3d) Static
    Dim length As Single
    length = Sqr(io.x * io.x + io.y * io.y + io.z * io.z)
    If length = 0.0 Then
        io.x = 0.0
        io.y = 0.0
        io.z = 0.0
    Else
        io.x = io.x / length
        io.y = io.y / length
        io.z = io.z / length
    End If
End Sub

Sub Vector3_NormalizeFlip (io As vec3d) Static
    ' normalize and also flip the direction
    Dim length As Single
    length = -Sqr(io.x * io.x + io.y * io.y + io.z * io.z)
    If length = 0.0 Then
        io.x = 0.0
        io.y = 0.0
        io.z = 0.0
    Else
        io.x = io.x / length
        io.y = io.y / length
        io.z = io.z / length
    End If
End Sub


Sub Vector3_Mul (left As vec3d, scale As Single, o As vec3d) Static
    o.x = left.x * scale
    o.y = left.y * scale
    o.z = left.z * scale
End Sub

Sub Vector3_CrossProduct (p0 As vec3d, p1 As vec3d, o As vec3d) Static
    o.x = p0.y * p1.z - p0.z * p1.y
    o.y = p0.z * p1.x - p0.x * p1.z
    o.z = p0.x * p1.y - p0.y * p1.x
End Sub

Sub CalcSurfaceNormal_3Point (p0 As vec3d, p1 As vec3d, p2 As vec3d, o As vec3d) Static
    Static line1_x As Single, line1_y As Single, line1_z As Single
    Static line2_x As Single, line2_y As Single, line2_z As Single
    Static lengthNormal As Single

    line1_x = p1.x - p0.x
    line1_y = p1.y - p0.y
    line1_z = p1.z - p0.z

    line2_x = p2.x - p0.x
    line2_y = p2.y - p0.y
    line2_z = p2.z - p0.z

    ' Cross Product
    o.x = line1_y * line2_z - line1_z * line2_y
    o.y = line1_z * line2_x - line1_x * line2_z
    o.z = line1_x * line2_y - line1_y * line2_x

    lengthNormal = Sqr(o.x * o.x + o.y * o.y + o.z * o.z)
    If lengthNormal = 0.0 Then Exit Sub

    o.x = o.x / lengthNormal
    o.y = o.y / lengthNormal
    o.z = o.z / lengthNormal

End Sub

Function Vector3_DotProduct! (p0 As vec3d, p1 As vec3d) Static
    Vector3_DotProduct! = p0.x * p1.x + p0.y * p1.y + p0.z * p1.z
End Function

Sub Vector3_Reflect (i As vec3d, normal As vec3d, o As vec3d) Static
    Static mag As Single
    Static bounce As vec3d
    mag = -2.0 * Vector3_DotProduct!(i, normal)
    Vector3_Mul normal, mag, bounce
    Vector3_Add bounce, i, o
End Sub

Sub Vector3_Reflect_unroll (i As vec3d, normal As vec3d, o As vec3d) Static
    Static mag As Single
    mag = -2.0 * (i.x * normal.x + i.y * normal.y + i.z * normal.z)
    o.x = i.x + normal.x * mag
    o.y = i.y + normal.y * mag
    o.z = i.z + normal.z * mag
End Sub

Sub ConvertXYZ_to_CubeIUV (x As Single, y As Single, z As Single, index As Integer, u As Single, v As Single)
    '     +---+
    '     | 2 |
    ' +---+---+---+---+
    ' | 1 | 4 | 0 | 5 |
    ' +---+---+---+---+
    '     | 3 |
    '     +---+
    Static absX As Single, absY As Single, absZ As Single
    absX = Abs(x)
    absY = Abs(y)
    absZ = Abs(z)

    If absX >= absY And absX >= absZ Then
        If x > 0 Then
            ' POSITIVE X
            ' u (0 to 1) goes from +z to -z
            ' v (0 to 1) goes from -y to +y
            index = 0
            ' Convert range from -1 to 1 to 0 to 1
            u = 0.5 * (-z / absX + 1.0)
            v = 0.5 * (-y / absX + 1.0)
            Exit Sub
        Else
            ' NEGATIVE X
            ' u (0 to 1) goes from -z to +z
            ' ?v (0 to 1) goes from -y to +y
            index = 1
            If absX = 0 Then
                ' Bail out
                u = 0.5
                v = 0.5
                Exit Sub
            End If
            ' Convert range from -1 to 1 to 0 to 1
            u = 0.5 * (z / absX + 1.0)
            v = 0.5 * (-y / absX + 1.0)
            Exit Sub
        End If
    End If

    If absY >= absX And absY >= absZ Then
        If y > 0 Then
            ' POSITIVE Y
            ' u (0 to 1) goes from -x to +x
            ' v (0 to 1) goes from +z to -z
            index = 2
            ' Convert range from -1 to 1 to 0 to 1
            u = 0.5 * (x / absY + 1.0)
            v = 0.5 * (z / absY + 1.0)
            Exit Sub
        Else
            ' NEGATIVE Y
            ' u (0 to 1) goes from -x to +x
            ' v (0 to 1) goes from -z to +z
            index = 3
            ' Convert range from -1 to 1 to 0 to 1
            u = 0.5 * (x / absY + 1.0)
            v = 0.5 * (-z / absY + 1.0)
            Exit Sub
        End If
    End If

    If z > 0 Then
        ' POSITIVE Z
        ' u (0 to 1) goes from -x to +x
        ' ?v (0 to 1) goes from -y to +y
        index = 4
        ' Convert range from -1 to 1 to 0 to 1
        u = 0.5 * (x / absZ + 1.0)
        v = 0.5 * (-y / absZ + 1.0)
        Exit Sub
    Else
        ' NEGATIVE Z
        ' u (0 to 1) goes from +x to -x
        ' v (0 to 1) goes from -y to +y
        index = 5
        ' Convert range from -1 to 1 to 0 to 1
        u = 0.5 * (-x / absZ + 1.0)
        v = 0.5 * (-y / absZ + 1.0)
    End If
End Sub

Sub Matrix4_MakeIdentity (m( 3 , 3) As Single)
    m(0, 0) = 1.0: m(0, 1) = 0.0: m(0, 2) = 0.0: m(0, 3) = 0.0
    m(1, 1) = 0.0: m(1, 1) = 1.0: m(1, 2) = 0.0: m(1, 3) = 0.0
    m(2, 2) = 0.0: m(2, 1) = 0.0: m(2, 2) = 1.0: m(2, 3) = 0.0
    m(3, 3) = 0.0: m(3, 1) = 0.0: m(3, 2) = 0.0: m(3, 3) = 1.0
End Sub

Sub Matrix4_MakeRotation_Z (deg As Single, m( 3 , 3) As Single)
    ' Rotation Z
    m(0, 0) = Cos(_D2R(deg))
    m(0, 1) = Sin(_D2R(deg))
    m(0, 2) = 0.0
    m(0, 3) = 0.0
    m(1, 0) = -Sin(_D2R(deg))
    m(1, 1) = Cos(_D2R(deg))
    m(1, 2) = 0.0
    m(1, 3) = 0.0
    m(2, 2) = 0.0: m(2, 1) = 0.0: m(2, 2) = 1.0: m(2, 3) = 0.0
    m(3, 3) = 0.0: m(3, 1) = 0.0: m(3, 2) = 0.0: m(3, 3) = 1.0
End Sub

Sub Matrix4_MakeRotation_Y (deg As Single, m( 3 , 3) As Single)
    m(0, 0) = Cos(_D2R(deg))
    m(0, 1) = 0.0
    m(0, 2) = Sin(_D2R(deg))
    m(0, 3) = 0.0
    m(1, 1) = 0.0: m(1, 1) = 1.0: m(1, 2) = 0.0: m(1, 3) = 0.0
    m(2, 0) = -Sin(_D2R(deg))
    m(2, 1) = 0.0
    m(2, 2) = Cos(_D2R(deg))
    m(2, 3) = 0.0
    m(3, 3) = 0.0: m(3, 1) = 0.0: m(3, 2) = 0.0: m(3, 3) = 1.0
End Sub

Sub Matrix4_MakeRotation_X (deg As Single, m( 3 , 3) As Single)
    m(0, 0) = 1.0: m(0, 1) = 0.0: m(0, 2) = 0.0: m(0, 3) = 0.0
    m(1, 0) = 0.0
    m(1, 1) = Cos(_D2R(deg))
    m(1, 2) = -Sin(_D2R(deg)) 'flip
    m(1, 3) = 0.0
    m(2, 0) = 0.0
    m(2, 1) = Sin(_D2R(deg)) 'flip
    m(2, 2) = Cos(_D2R(deg))
    m(2, 3) = 0.0
    m(3, 3) = 0.0: m(3, 1) = 0.0: m(3, 2) = 0.0: m(3, 3) = 1.0
End Sub

Sub Matrix4_PointAt (psn As vec3d, target As vec3d, up As vec3d, m( 3 , 3) As Single)
    ' Calculate new forward direction
    Dim newForward As vec3d
    Vector3_Delta target, psn, newForward
    Vector3_Normalize newForward

    ' Calculate new Up direction
    Dim a As vec3d
    Dim newUp As vec3d
    Vector3_Mul newForward, Vector3_DotProduct(up, newForward), a
    Vector3_Delta up, a, newUp
    Vector3_Normalize newUp

    ' new Right direction is just cross product
    Dim newRight As vec3d
    Vector3_CrossProduct newUp, newForward, newRight

    ' Construct Dimensioning and Translation Matrix
    m(0, 0) = newRight.x: m(0, 1) = newRight.y: m(0, 2) = newRight.z: m(0, 3) = 0.0
    m(1, 0) = newUp.x: m(1, 1) = newUp.y: m(1, 2) = newUp.z: m(1, 3) = 0.0
    m(2, 0) = newForward.x: m(2, 1) = newForward.y: m(2, 2) = newForward.z: m(2, 3) = 0.0
    m(3, 0) = psn.x: m(3, 1) = psn.y: m(3, 2) = psn.z: m(3, 3) = 1.0

End Sub

Sub Matrix4_QuickInverse (m( 3 , 3) As Single, q( 3 , 3) As Single)
    q(0, 0) = m(0, 0): q(0, 1) = m(1, 0): q(0, 2) = m(2, 0): q(0, 3) = 0.0
    q(1, 0) = m(0, 1): q(1, 1) = m(1, 1): q(1, 2) = m(2, 1): q(1, 3) = 0.0
    q(2, 0) = m(0, 2): q(2, 1) = m(1, 2): q(2, 2) = m(2, 2): q(2, 3) = 0.0
    q(3, 0) = -(m(3, 0) * q(0, 0) + m(3, 1) * q(1, 0) + m(3, 2) * q(2, 0))
    q(3, 1) = -(m(3, 0) * q(0, 1) + m(3, 1) * q(1, 1) + m(3, 2) * q(2, 1))
    q(3, 2) = -(m(3, 0) * q(0, 2) + m(3, 1) * q(1, 2) + m(3, 2) * q(2, 2))
    q(3, 3) = 1.0
End Sub

' Projection is another optimized matrix multiplcation.
' w is assumed 1 on the input side, but this time it is necessary to output.
' X and Y needs to be normalized.
' Z is unused.
Sub ProjectMatrixVector4 (i As vec3d, m( 3 , 3) As Single, o As vec4d)
    Dim www As Single
    o.x = i.x * m(0, 0) + i.y * m(1, 0) + i.z * m(2, 0) + m(3, 0)
    o.y = i.x * m(0, 1) + i.y * m(1, 1) + i.z * m(2, 1) + m(3, 1)
    o.z = i.x * m(0, 2) + i.y * m(1, 2) + i.z * m(2, 2) + m(3, 2)
    www = i.x * m(0, 3) + i.y * m(1, 3) + i.z * m(2, 3) + m(3, 3)

    ' Normalizing
    If www <> 0.0 Then
        o.w = 1 / www 'optimization
        o.x = o.x * o.w
        o.y = -o.y * o.w 'because I feel +Y is up
        o.z = o.z * o.w
    End If
End Sub

Sub QuickSort_ViewZ (start As Long, finish As Long, J() As objectlist_type)
    Dim Lo As Long
    Dim Hi As Long
    Dim middleVal As Single

    Lo = start
    Hi = finish
    middleVal = J((Lo + Hi) / 2).viewz
    Do
        Do While J(Lo).viewz < middleVal: Lo = Lo + 1: Loop
        Do While J(Hi).viewz > middleVal: Hi = Hi - 1: Loop
        If Lo <= Hi Then
            Swap J(Lo), J(Hi)
            Lo = Lo + 1: Hi = Hi - 1
        End If
    Loop Until Lo > Hi
    If Hi > start Then Call QuickSort_ViewZ(start, Hi, J())
    If Lo < finish Then Call QuickSort_ViewZ(Lo, finish, J())
End Sub


Function ReadTexel& (ccol As Single, rrow As Single) Static
    Select Case T1_Filter_Selection
        Case 0
            ReadTexel& = ReadTexelNearest&(ccol, rrow)
        Case 1
            ReadTexel& = ReadTexel3Point&(ccol, rrow)
        Case 2
            ReadTexel& = ReadTexelBiLinearFix&(ccol, rrow)
        Case 3
            ReadTexel& = ReadTexelBiLinear&(ccol, rrow)
    End Select
End Function


Function ReadTexelNearest& (ccol As Single, rrow As Single) Static
    ' Relies on some shared variables over by Texture1
    Static cc As Integer
    Static rr As Integer

    ' Decided just to tile the texture if out of bounds
    cc = Int(ccol) And T1_width_MASK
    rr = Int(rrow) And T1_height_MASK

    ReadTexelNearest& = Texture1(cc, rr)
End Function


Function ReadTexel3Point& (ccol As Single, rrow As Single) Static
    ' Relies on some shared T1 variables over by Texture1
    Static cc As Integer
    Static rr As Integer
    Static cc1 As Integer
    Static rr1 As Integer

    Static Frac_cc1 As Single
    Static Frac_rr1 As Single

    Static Area_00 As Single
    Static Area_11 As Single
    Static Area_2f As Single

    Static uv_0_0 As Long
    Static uv_1_1 As Long
    Static uv_f As Long

    Static r0 As Long
    Static g0 As Long
    Static b0 As Long

    Static cm5 As Single
    Static rm5 As Single

    ' Offset so the transition appears in the center of an enlarged texel instead of a corner.
    cm5 = ccol - 0.5
    rm5 = rrow - 0.5

    If T1_options And T1_option_clamp_width Then
        ' clamp
        If cm5 < 0.0 Then cm5 = 0.0
        If cm5 >= T1_width_MASK Then
            ' 15.0 and up
            cc = T1_width_MASK
            cc1 = T1_width_MASK
        Else
            ' 0 1 2 .. 13 14.999
            cc = Int(cm5)
            cc1 = cc + 1
        End If
    Else
        ' tile the texture
        cc = Int(cm5) And T1_width_MASK
        cc1 = (cc + 1) And T1_width_MASK
    End If

    If T1_options And T1_option_clamp_height Then
        ' clamp
        If rm5 < 0.0 Then rm5 = 0.0
        If rm5 >= T1_height_MASK Then
            rr = T1_height_MASK
            rr1 = T1_height_MASK
        Else
            rr = Int(rm5)
            rr1 = rr + 1
        End If
    Else
        ' tile
        rr = Int(rm5) And T1_height_MASK
        rr1 = (rr + 1) And T1_height_MASK
    End If

    uv_0_0 = Texture1(cc, rr)
    uv_1_1 = Texture1(cc1, rr1)

    Frac_cc1 = cm5 - Int(cm5)
    Frac_rr1 = rm5 - Int(rm5)

    If Frac_cc1 > Frac_rr1 Then
        ' top-right
        ' Area of a triangle = 1/2 * base * height
        ' Using twice the areas (rectangles) to eliminate a multiply by 1/2 and a later divide by 1/2
        Area_11 = Frac_rr1
        Area_00 = 1.0 - Frac_cc1
        uv_f = Texture1(cc1, rr)
    Else
        ' bottom-left
        Area_00 = 1.0 - Frac_rr1
        Area_11 = Frac_cc1
        uv_f = Texture1(cc, rr1)
    End If

    Area_2f = 1.0 - (Area_00 + Area_11) '1.0 here is twice the total triangle area.

    r0 = _Red32(uv_f) * Area_2f + _Red32(uv_0_0) * Area_00 + _Red32(uv_1_1) * Area_11
    g0 = _Green32(uv_f) * Area_2f + _Green32(uv_0_0) * Area_00 + _Green32(uv_1_1) * Area_11
    b0 = _Blue32(uv_f) * Area_2f + _Blue32(uv_0_0) * Area_00 + _Blue32(uv_1_1) * Area_11

    'ReadTexel3Point& = _RGB32(r0, g0, b0)
    ReadTexel3Point& = _ShL(r0, 16) Or _ShL(g0, 8) Or b0
End Function


Function ReadTexelBiLinear& (ccol As Single, rrow As Single) Static
    ' Relies on some shared variables over by Texture1
    Static cc As Integer
    Static rr As Integer
    Static cc1 As Integer
    Static rr1 As Integer

    Static frac_cc1 As Single
    Static frac_rr1 As Single

    ' caching of 4 texels
    Static this_cache As _Unsigned Long
    Static uv_0_0 As Long
    Static uv_0_1 As Long
    Static uv_1_0 As Long
    Static uv_1_1 As Long

    Static r0 As Long
    Static g0 As Long
    Static b0 As Long
    Static r1 As Long
    Static g1 As Long
    Static b1 As Long

    Static cm5 As Single
    Static rm5 As Single

    cm5 = ccol - 0.5
    rm5 = rrow - 0.5

    If T1_options And T1_option_clamp_width Then
        ' clamp
        If cm5 < 0.0 Then cm5 = 0.0
        If cm5 >= T1_width_MASK Then
            ' 15.0 and up
            cc = T1_width_MASK
            cc1 = T1_width_MASK
        Else
            ' 0 1 2 .. 13 14.999
            cc = Int(cm5)
            cc1 = cc + 1
        End If
    Else
        ' tile the texture
        cc = Int(cm5) And T1_width_MASK
        cc1 = (cc + 1) And T1_width_MASK
    End If

    If T1_options And T1_option_clamp_height Then
        ' clamp
        If rm5 < 0.0 Then rm5 = 0.0
        If rm5 >= T1_height_MASK Then
            rr = T1_height_MASK
            rr1 = T1_height_MASK
        Else
            rr = Int(rm5)
            rr1 = rr + 1
        End If
    Else
        ' tile
        rr = Int(rm5) And T1_height_MASK
        rr1 = (rr + 1) And T1_height_MASK
    End If

    uv_0_0 = Texture1(cc, rr)
    uv_1_0 = Texture1(cc1, rr)
    uv_0_1 = Texture1(cc, rr1)
    uv_1_1 = Texture1(cc1, rr1)

    this_cache = _ShL(rr, 16) Or cc
    If this_cache <> T1_last_cache Then
        uv_0_0 = Texture1(cc, rr)
        uv_1_0 = Texture1(cc1, rr)
        uv_0_1 = Texture1(cc, rr1)
        uv_1_1 = Texture1(cc1, rr1)
        T1_last_cache = this_cache
        ' uncomment below to show cache miss in yellow
        'ReadTexelBiLinearFix& = _RGB32(255, 255, 127)
        'Exit Function
    End If

    frac_cc1 = cm5 - Int(cm5)
    frac_rr1 = rm5 - Int(rm5)

    r0 = _Red32(uv_0_0)
    r0 = (_Red32(uv_1_0) - r0) * frac_cc1 + r0

    g0 = _Green32(uv_0_0)
    g0 = (_Green32(uv_1_0) - g0) * frac_cc1 + g0

    b0 = _Blue32(uv_0_0)
    b0 = (_Blue32(uv_1_0) - b0) * frac_cc1 + b0

    r1 = _Red32(uv_0_1)
    r1 = (_Red32(uv_1_1) - r1) * frac_cc1 + r1

    g1 = _Green32(uv_0_1)
    g1 = (_Green32(uv_1_1) - g1) * frac_cc1 + g1

    b1 = _Blue32(uv_0_1)
    b1 = (_Blue32(uv_1_1) - b1) * frac_cc1 + b1

    ReadTexelBiLinear& = _RGB32((r1 - r0) * frac_rr1 + r0, (g1 - g0) * frac_rr1 + g0, (b1 - b0) * frac_rr1 + b0)
End Function


Function ReadTexelBiLinearFix& (ccol As Single, rrow As Single) Static
    ' caching of 4 texels
    Static this_cache As _Unsigned Long
    Static uv_0_0 As Long
    Static uv_0_1 As Long
    Static uv_1_0 As Long
    Static uv_1_1 As Long

    ' Relies on some shared variables over by Texture1
    Static cc As _Unsigned Integer
    Static rr As _Unsigned Integer
    Static cc1 As _Unsigned Integer
    Static rr1 As _Unsigned Integer

    'Static Frac_cc1 As Single
    'Static Frac_rr1 As Single

    Static Frac_cc1_FIX7 As Integer
    Static Frac_rr1_FIX7 As Integer

    Static r0 As Integer
    Static g0 As Integer
    Static b0 As Integer
    Static r1 As Integer
    Static g1 As Integer
    Static b1 As Integer

    Static cm5 As Single
    Static rm5 As Single

    cm5 = ccol - 0.5
    rm5 = rrow - 0.5

    If T1_options And T1_option_clamp_width Then
        ' clamp
        If cm5 < 0.0 Then cm5 = 0.0
        If cm5 >= T1_width_MASK Then
            ' 15.0 and up
            cc = T1_width_MASK
            cc1 = T1_width_MASK
        Else
            ' 0 1 2 .. 13 14.999
            cc = Int(cm5)
            cc1 = cc + 1
        End If
    Else
        ' tile the texture
        cc = Int(cm5) And T1_width_MASK
        cc1 = (cc + 1) And T1_width_MASK
    End If

    If T1_options And T1_option_clamp_height Then
        ' clamp
        If rm5 < 0.0 Then rm5 = 0.0
        If rm5 >= T1_height_MASK Then
            rr = T1_height_MASK
            rr1 = T1_height_MASK
        Else
            rr = Int(rm5)
            rr1 = rr + 1
        End If
    Else
        ' tile
        rr = Int(rm5) And T1_height_MASK
        rr1 = (rr + 1) And T1_height_MASK
    End If

    'Frac_cc1 = cm5 - Int(cm5)
    Frac_cc1_FIX7 = (cm5 - Int(cm5)) * 128
    'Frac_rr1 = rm5 - Int(rm5)
    Frac_rr1_FIX7 = (rm5 - Int(rm5)) * 128

    ' cache
    this_cache = _ShL(rr, 16) Or cc
    If this_cache <> T1_last_cache Then
        uv_0_0 = Texture1(cc, rr)
        uv_1_0 = Texture1(cc1, rr)
        uv_0_1 = Texture1(cc, rr1)
        uv_1_1 = Texture1(cc1, rr1)
        T1_last_cache = this_cache
        ' uncomment below to show cache miss in yellow
        'ReadTexelBiLinearFix& = _RGB32(255, 255, 127)
        'Exit Function
    End If

    r0 = _Red32(uv_0_0)
    r0 = _ShR((_Red32(uv_1_0) - r0) * Frac_cc1_FIX7, 7) + r0

    g0 = _Green32(uv_0_0)
    g0 = _ShR((_Green32(uv_1_0) - g0) * Frac_cc1_FIX7, 7) + g0

    b0 = _Blue32(uv_0_0)
    b0 = _ShR((_Blue32(uv_1_0) - b0) * Frac_cc1_FIX7, 7) + b0

    r1 = _Red32(uv_0_1)
    r1 = _ShR((_Red32(uv_1_1) - r1) * Frac_cc1_FIX7, 7) + r1

    g1 = _Green32(uv_0_1)
    g1 = _ShR((_Green32(uv_1_1) - g1) * Frac_cc1_FIX7, 7) + g1

    b1 = _Blue32(uv_0_1)
    b1 = _ShR((_Blue32(uv_1_1) - b1) * Frac_cc1_FIX7, 7) + b1

    ReadTexelBiLinearFix& = _ShL(_ShR((r1 - r0) * Frac_rr1_FIX7, 7) + r0, 16) Or _ShL(_ShR((g1 - g0) * Frac_rr1_FIX7, 7) + g0, 8) Or _ShR((b1 - b0) * Frac_rr1_FIX7, 7) + b0
End Function


Function RGB_Fog& (zz As Single, RGB_color As _Unsigned Long) Static
    Static r0 As Long
    Static g0 As Long
    Static b0 As Long
    Static fog_scale As Single

    If zz <= Fog_near Then
        RGB_Fog& = RGB_color
    ElseIf zz >= Fog_far Then
        RGB_Fog& = Fog_color
    Else
        fog_scale = (zz - Fog_near) * Fog_rate
        r0 = _Red32(RGB_color)
        g0 = _Green32(RGB_color)
        b0 = _Blue32(RGB_color)

        RGB_Fog& = _RGB32((Fog_R - r0) * fog_scale + r0, (Fog_G - g0) * fog_scale + g0, (Fog_B - b0) * fog_scale + b0)
    End If
End Function


Function RGB_Lit& (RGB_color As _Unsigned Long) Static
    Static scale As Single
    scale = Light_Directional + Light_AmbientVal ' oversaturate the bright colors

    RGB_Lit& = _RGB32(scale * _Red32(RGB_color), scale * _Green32(RGB_color), scale * _Blue32(RGB_color)) 'values over 255 are just clamped to 255
End Function


Function RGB_Sum& (RGB_1 As _Unsigned Long, RGB_2 As _Unsigned Long) Static
    ' Lighten function
    Static r1 As Long
    Static g1 As Long
    Static b1 As Long
    Static r2 As Long
    Static g2 As Long
    Static b2 As Long

    r1 = _Red32(RGB_1)
    g1 = _Green32(RGB_1)
    b1 = _Blue32(RGB_1)
    r2 = _Red32(RGB_2)
    g2 = _Green32(RGB_2)
    b2 = _Blue32(RGB_2)

    RGB_Sum& = _RGB32(r1 + r2, g1 + g2, b1 + b2) ' values over 255 are just clamped to 255
End Function


Function RGB_Modulate& (RGB_1 As _Unsigned Long, RGB_Mod As _Unsigned Long) Static
    ' Darken function
    Static r1 As Integer
    Static g1 As Integer
    Static b1 As Integer
    Static r2 As Integer
    Static g2 As Integer
    Static b2 As Integer

    r1 = _Red32(RGB_1)
    g1 = _Green32(RGB_1)
    b1 = _Blue32(RGB_1)
    r2 = _Red32(RGB_Mod) + 1 'make it so that 255 is 1.0
    g2 = _Green32(RGB_Mod) + 1 'but do it in a way that is not a division
    b2 = _Blue32(RGB_Mod) + 1

    RGB_Modulate& = _RGB32(_ShR(r1 * r2, 8), _ShR(g1 * g2, 8), _ShR(b1 * b2, 8))
End Function

Sub TexturedVertexColorAlphaTriangle (A As vertex9, B As vertex9, C As vertex9)
    Static delta2 As vertex9
    Static delta1 As vertex9
    Static draw_min_y As Long, draw_max_y As Long

    ' Sort so that vertex A is on top and C is on bottom.
    ' This seems inverted from math class, but remember that Y increases in value downward on PC monitors
    If B.y < A.y Then
        Swap A, B
    End If
    If C.y < A.y Then
        Swap A, C
    End If
    If C.y < B.y Then
        Swap B, C
    End If

    ' integer window clipping
    draw_min_y = _Ceil(A.y)
    If draw_min_y < clip_min_y Then draw_min_y = clip_min_y
    draw_max_y = _Ceil(C.y) - 1
    If draw_max_y > clip_max_y Then draw_max_y = clip_max_y
    If (draw_max_y - draw_min_y) < 0 Then Exit Sub

    ' Determine the deltas (lengths)
    ' delta 2 is from A to C (the full triangle height)
    delta2.x = C.x - A.x
    delta2.y = C.y - A.y
    delta2.w = C.w - A.w
    delta2.u = C.u - A.u
    delta2.v = C.v - A.v
    delta2.r = C.r - A.r
    delta2.g = C.g - A.g
    delta2.b = C.b - A.b
    delta2.a = C.a - A.a

    ' Avoiding div by 0
    ' Entire Y height less than 1/256 would not have meaningful pixel color change
    If delta2.y < (1 / 256) Then Exit Sub

    ' Determine vertical Y steps for DDA style math
    ' DDA is Digital Differential Analyzer
    ' It is an accumulator that counts from a known start point to an end point, in equal increments defined by the number of steps in-between.
    ' Probably faster nowadays to do the one division at the start, instead of Bresenham, anyway.
    Static legx1_step As Single
    Static legw1_step As Single, legu1_step As Single, legv1_step As Single
    Static legr1_step As Single, legg1_step As Single, legb1_step As Single
    Static lega1_step As Single

    Static legx2_step As Single
    Static legw2_step As Single, legu2_step As Single, legv2_step As Single
    Static legr2_step As Single, legg2_step As Single, legb2_step As Single
    Static lega2_step As Single

    ' Leg 2 steps from A to C (the full triangle height)
    legx2_step = delta2.x / delta2.y
    legw2_step = delta2.w / delta2.y
    legu2_step = delta2.u / delta2.y
    legv2_step = delta2.v / delta2.y
    legr2_step = delta2.r / delta2.y
    legg2_step = delta2.g / delta2.y
    legb2_step = delta2.b / delta2.y
    lega2_step = delta2.a / delta2.y

    ' Leg 1, Draw top to middle
    ' For most triangles, draw downward from the apex A to a knee B.
    ' That knee could be on either the left or right side, but that is handled much later.
    Static draw_middle_y As Long
    draw_middle_y = _Ceil(B.y)
    If draw_middle_y < clip_min_y Then draw_middle_y = clip_min_y
    ' Do not clip B to max_y. Let the y count expire before reaching the knee if it is past bottom of screen.

    ' Leg 1 is from A to B (right now)
    delta1.x = B.x - A.x
    delta1.y = B.y - A.y
    delta1.w = B.w - A.w
    delta1.u = B.u - A.u
    delta1.v = B.v - A.v
    delta1.r = B.r - A.r
    delta1.g = B.g - A.g
    delta1.b = B.b - A.b
    delta1.a = B.a - A.a

    ' If the triangle has no knee, this section gets skipped to avoid divide by 0.
    ' That is okay, because the recalculate Leg 1 from B to C triggers before actually drawing.
    If delta1.y > (1 / 256) Then
        ' Find Leg 1 steps in the y direction from A to B
        legx1_step = delta1.x / delta1.y
        legw1_step = delta1.w / delta1.y
        legu1_step = delta1.u / delta1.y
        legv1_step = delta1.v / delta1.y
        legr1_step = delta1.r / delta1.y
        legg1_step = delta1.g / delta1.y
        legb1_step = delta1.b / delta1.y
        lega1_step = delta1.a / delta1.y
    End If

    ' Y Accumulators
    Static leg_x1 As Single
    Static leg_w1 As Single, leg_u1 As Single, leg_v1 As Single
    Static leg_r1 As Single, leg_g1 As Single, leg_b1 As Single
    Static leg_a1 As Single

    Static leg_x2 As Single
    Static leg_w2 As Single, leg_u2 As Single, leg_v2 As Single
    Static leg_r2 As Single, leg_g2 As Single, leg_b2 As Single
    Static leg_a2 As Single

    ' 11-4-2022 Prestep Y
    Static prestep_y1 As Single
    ' Basically we are sampling pixels on integer exact rows.
    ' But we only are able to know the next row by way of forward interpolation. So always round up.
    ' To get to that next row, we have to prestep by the fractional forward distance from A. _Ceil(A.y) - A.y
    prestep_y1 = draw_min_y - A.y

    leg_x1 = A.x + prestep_y1 * legx1_step
    leg_w1 = A.w + prestep_y1 * legw1_step
    leg_u1 = A.u + prestep_y1 * legu1_step
    leg_v1 = A.v + prestep_y1 * legv1_step
    leg_r1 = A.r + prestep_y1 * legr1_step
    leg_g1 = A.g + prestep_y1 * legg1_step
    leg_b1 = A.b + prestep_y1 * legb1_step
    leg_a1 = A.a + prestep_y1 * lega1_step

    leg_x2 = A.x + prestep_y1 * legx2_step
    leg_w2 = A.w + prestep_y1 * legw2_step
    leg_u2 = A.u + prestep_y1 * legu2_step
    leg_v2 = A.v + prestep_y1 * legv2_step
    leg_r2 = A.r + prestep_y1 * legr2_step
    leg_g2 = A.g + prestep_y1 * legg2_step
    leg_b2 = A.b + prestep_y1 * legb2_step
    leg_a2 = A.a + prestep_y1 * lega2_step

    ' Inner loop vars
    Static row As Long
    Static col As Long
    Static draw_max_x As Long
    Static zbuf_index As _Unsigned Long ' Z-Buffer
    Static tex_z As Single ' 1/w helper (multiply by inverse is faster than dividing each time)
    Static pixel_alpha As Single

    ' Stepping along the X direction
    Static delta_x As Single
    Static prestep_x As Single
    Static tex_w_step As Single, tex_u_step As Single, tex_v_step As Single
    Static tex_r_step As Single, tex_g_step As Single, tex_b_step As Single
    Static tex_a_step As Single

    ' X Accumulators
    Static tex_w As Single, tex_u As Single, tex_v As Single
    Static tex_r As Single, tex_g As Single, tex_b As Single
    Static tex_a As Single

    ' Work Screen Memory Pointers
    Static work_mem_info As _MEM
    Static work_next_row_step As _Offset
    Static work_row_base As _Offset ' Calculated every row
    Static work_address As _Offset ' Calculated at every starting column
    work_mem_info = _MemImage(WORK_IMAGE)
    work_next_row_step = 4 * Size_Render_X

    ' Invalidate texel cache
    T1_last_cache = &HFFFFFFFF

    ' Row Loop from top to bottom
    row = draw_min_y
    work_row_base = work_mem_info.OFFSET + row * work_next_row_step
    While row <= draw_max_y

        If row = draw_middle_y Then
            ' Reached Leg 1 knee at B, recalculate Leg 1.
            ' This overwrites Leg 1 to be from B to C. Leg 2 just keeps continuing from A to C.
            delta1.x = C.x - B.x
            delta1.y = C.y - B.y
            delta1.w = C.w - B.w
            delta1.u = C.u - B.u
            delta1.v = C.v - B.v
            delta1.r = C.r - B.r
            delta1.g = C.g - B.g
            delta1.b = C.b - B.b
            delta1.a = C.a - B.a

            If delta1.y = 0.0 Then Exit Sub

            ' Full steps in the y direction from B to C
            legx1_step = delta1.x / delta1.y
            legw1_step = delta1.w / delta1.y
            legu1_step = delta1.u / delta1.y
            legv1_step = delta1.v / delta1.y
            legr1_step = delta1.r / delta1.y ' vertex color
            legg1_step = delta1.g / delta1.y
            legb1_step = delta1.b / delta1.y
            lega1_step = delta1.a / delta1.y

            ' 11-4-2022 Prestep Y
            ' Most cases has B lower downscreen than A.
            ' B > A usually. Only one case where B = A.
            prestep_y1 = draw_middle_y - B.y

            ' Re-Initialize DDA start values
            leg_x1 = B.x + prestep_y1 * legx1_step
            leg_w1 = B.w + prestep_y1 * legw1_step
            leg_u1 = B.u + prestep_y1 * legu1_step
            leg_v1 = B.v + prestep_y1 * legv1_step
            leg_r1 = B.r + prestep_y1 * legr1_step
            leg_g1 = B.g + prestep_y1 * legg1_step
            leg_b1 = B.b + prestep_y1 * legb1_step
            leg_a1 = B.a + prestep_y1 * lega1_step
        End If

        ' Horizontal Scanline
        delta_x = Abs(leg_x2 - leg_x1)
        ' Avoid div/0, this gets tiring.
        If delta_x >= (1 / 2048) Then
            ' Calculate step, start, and end values.
            ' Drawing left to right, as in incrementing from a lower to higher memory address, is usually fastest.
            If leg_x1 < leg_x2 Then
                ' leg 1 is on the left
                tex_w_step = (leg_w2 - leg_w1) / delta_x
                tex_u_step = (leg_u2 - leg_u1) / delta_x
                tex_v_step = (leg_v2 - leg_v1) / delta_x
                tex_r_step = (leg_r2 - leg_r1) / delta_x
                tex_g_step = (leg_g2 - leg_g1) / delta_x
                tex_b_step = (leg_b2 - leg_b1) / delta_x
                tex_a_step = (leg_a2 - leg_a1) / delta_x

                ' Set the horizontal starting point to (1)
                col = _Ceil(leg_x1)
                If col < clip_min_x Then col = clip_min_x

                ' Prestep to find pixel starting point
                prestep_x = col - leg_x1
                tex_w = leg_w1 + prestep_x * tex_w_step
                tex_z = 1 / tex_w ' this can be absorbed
                tex_u = leg_u1 + prestep_x * tex_u_step
                tex_v = leg_v1 + prestep_x * tex_v_step
                tex_r = leg_r1 + prestep_x * tex_r_step
                tex_g = leg_g1 + prestep_x * tex_g_step
                tex_b = leg_b1 + prestep_x * tex_b_step
                tex_a = leg_a1 + prestep_x * tex_a_step

                ' ending point is (2)
                draw_max_x = _Ceil(leg_x2)
                If draw_max_x > clip_max_x Then draw_max_x = clip_max_x

            Else
                ' Things are flipped. leg 1 is on the right.
                tex_w_step = (leg_w1 - leg_w2) / delta_x
                tex_u_step = (leg_u1 - leg_u2) / delta_x
                tex_v_step = (leg_v1 - leg_v2) / delta_x
                tex_r_step = (leg_r1 - leg_r2) / delta_x
                tex_g_step = (leg_g1 - leg_g2) / delta_x
                tex_b_step = (leg_b1 - leg_b2) / delta_x
                tex_a_step = (leg_a1 - leg_a2) / delta_x

                ' Set the horizontal starting point to (2)
                col = _Ceil(leg_x2)
                If col < clip_min_x Then col = clip_min_x

                ' Prestep to find pixel starting point
                prestep_x = col - leg_x2
                tex_w = leg_w2 + prestep_x * tex_w_step
                tex_z = 1 / tex_w ' this can be absorbed
                tex_u = leg_u2 + prestep_x * tex_u_step
                tex_v = leg_v2 + prestep_x * tex_v_step
                tex_r = leg_r2 + prestep_x * tex_r_step
                tex_g = leg_g2 + prestep_x * tex_g_step
                tex_b = leg_b2 + prestep_x * tex_b_step
                tex_a = leg_a2 + prestep_x * tex_a_step

                ' ending point is (1)
                draw_max_x = _Ceil(leg_x1)
                If draw_max_x > clip_max_x Then draw_max_x = clip_max_x

            End If

            ' Draw the Horizontal Scanline
            ' Optimization: before entering this loop, must have done tex_z = 1 / tex_w
            work_address = work_row_base + 4 * col
            zbuf_index = row * Size_Render_X + col
            While col < draw_max_x

                If Screen_Z_Buffer(zbuf_index) = 0.0 Or tex_z < Screen_Z_Buffer(zbuf_index) Then
                    If (T1_options And T1_option_no_Z_write) = 0 Then
                        Screen_Z_Buffer(zbuf_index) = tex_z + Z_Fight_Bias
                    End If

                    Static pixel_combine As _Unsigned Long
                    If (T1_option_no_T1 And T1_options) Then
                        pixel_combine = _RGB32(tex_r, tex_g, tex_b)
                    Else
                        pixel_combine = RGB_Modulate(ReadTexel&(tex_u * tex_z, tex_v * tex_z), _RGB32(tex_r, tex_g, tex_b))
                    End If

                    Static pixel_existing As _Unsigned Long
                    pixel_alpha = tex_a * tex_z
                    If pixel_alpha < 0.998 Then
                        pixel_existing = _MemGet(work_mem_info, work_address, _Unsigned Long)
                        pixel_combine = _RGB32((  _red32(pixel_combine) - _Red32(pixel_existing))   * pixel_alpha + _red32(pixel_existing), _
                                               (_green32(pixel_combine) - _Green32(pixel_existing)) * pixel_alpha + _green32(pixel_existing), _
                                               ( _Blue32(pixel_combine) - _Blue32(pixel_existing))  * pixel_alpha + _blue32(pixel_existing))

                        ' x = (p1 - p0) * ratio + p0 is equivalent to
                        ' x = (1.0 - ratio) * p0 + ratio * p1
                    End If
                    _MemPut work_mem_info, work_address, pixel_combine
                    'If Dither_Selection > 0 Then
                    '    RGB_Dither555 col, row, pixel_value
                    'Else
                    '    PSet (col, row), pixel_value
                    'End If

                End If ' tex_z
                zbuf_index = zbuf_index + 1
                tex_w = tex_w + tex_w_step
                tex_z = 1 / tex_w ' floating point divide can be done in parallel when result not required immediately.
                tex_u = tex_u + tex_u_step
                tex_v = tex_v + tex_v_step
                tex_r = tex_r + tex_r_step
                tex_g = tex_g + tex_g_step
                tex_b = tex_b + tex_b_step
                tex_a = tex_a + tex_a_step
                work_address = work_address + 4
                col = col + 1
            Wend ' col

        End If ' end div/0 avoidance

        ' DDA next step
        leg_x1 = leg_x1 + legx1_step
        leg_w1 = leg_w1 + legw1_step
        leg_u1 = leg_u1 + legu1_step
        leg_v1 = leg_v1 + legv1_step
        leg_r1 = leg_r1 + legr1_step
        leg_g1 = leg_g1 + legg1_step
        leg_b1 = leg_b1 + legb1_step
        leg_a1 = leg_a1 + lega1_step

        leg_x2 = leg_x2 + legx2_step
        leg_w2 = leg_w2 + legw2_step
        leg_u2 = leg_u2 + legu2_step
        leg_v2 = leg_v2 + legv2_step
        leg_r2 = leg_r2 + legr2_step
        leg_g2 = leg_g2 + legg2_step
        leg_b2 = leg_b2 + legb2_step
        leg_a2 = leg_a2 + lega2_step

        work_row_base = work_row_base + work_next_row_step
        row = row + 1
    Wend ' row

End Sub

Sub TexturedNonlitTriangle (A As vertex9, B As vertex9, C As vertex9)
    ' this is a reduced copy for skybox drawing
    ' Texture_options is ignored
    Static delta2 As vertex5
    Static delta1 As vertex5
    Static draw_min_y As Long, draw_max_y As Long

    ' Sort so that vertex A is on top and C is on bottom.
    ' This seems inverted from math class, but remember that Y increases in value downward on PC monitors
    If B.y < A.y Then
        Swap A, B
    End If
    If C.y < A.y Then
        Swap A, C
    End If
    If C.y < B.y Then
        Swap B, C
    End If

    ' integer window clipping
    draw_min_y = _Ceil(A.y)
    If draw_min_y < clip_min_y Then draw_min_y = clip_min_y
    draw_max_y = _Ceil(C.y) - 1
    If draw_max_y > clip_max_y Then draw_max_y = clip_max_y
    If (draw_max_y - draw_min_y) < 0 Then Exit Sub

    ' Determine the deltas (lengths)
    ' delta 2 is from A to C (the full triangle height)
    delta2.x = C.x - A.x
    delta2.y = C.y - A.y
    delta2.w = C.w - A.w
    delta2.u = C.u - A.u
    delta2.v = C.v - A.v

    ' Avoiding div by 0
    ' Entire Y height less than 1/256 would not have meaningful pixel color change
    If delta2.y < (1 / 256) Then Exit Sub

    ' Determine vertical Y steps for DDA style math
    ' DDA is Digital Differential Analyzer
    ' It is an accumulator that counts from a known start point to an end point, in equal increments defined by the number of steps in-between.
    ' Probably faster nowadays to do the one division at the start, instead of Bresenham, anyway.
    Static legx1_step As Single
    Static legw1_step As Single, legu1_step As Single, legv1_step As Single

    Static legx2_step As Single
    Static legw2_step As Single, legu2_step As Single, legv2_step As Single

    ' Leg 2 steps from A to C (the full triangle height)
    legx2_step = delta2.x / delta2.y
    legw2_step = delta2.w / delta2.y
    legu2_step = delta2.u / delta2.y
    legv2_step = delta2.v / delta2.y

    ' Leg 1, Draw top to middle
    ' For most triangles, draw downward from the apex A to a knee B.
    ' That knee could be on either the left or right side, but that is handled much later.
    Static draw_middle_y As Long
    draw_middle_y = _Ceil(B.y)
    If draw_middle_y < clip_min_y Then draw_middle_y = clip_min_y
    ' Do not clip B to max_y. Let the y count expire before reaching the knee if it is past bottom of screen.

    ' Leg 1 is from A to B (right now)
    delta1.x = B.x - A.x
    delta1.y = B.y - A.y
    delta1.w = B.w - A.w
    delta1.u = B.u - A.u
    delta1.v = B.v - A.v

    ' If the triangle has no knee, this section gets skipped to avoid divide by 0.
    ' That is okay, because the recalculate Leg 1 from B to C triggers before actually drawing.
    If delta1.y > (1 / 256) Then
        ' Find Leg 1 steps in the y direction from A to B
        legx1_step = delta1.x / delta1.y
        legw1_step = delta1.w / delta1.y
        legu1_step = delta1.u / delta1.y
        legv1_step = delta1.v / delta1.y
    End If

    ' Y Accumulators
    Static leg_x1 As Single
    Static leg_w1 As Single, leg_u1 As Single, leg_v1 As Single

    Static leg_x2 As Single
    Static leg_w2 As Single, leg_u2 As Single, leg_v2 As Single

    ' 11-4-2022 Prestep Y
    Static prestep_y1 As Single
    ' Basically we are sampling pixels on integer exact rows.
    ' But we only are able to know the next row by way of forward interpolation. So always round up.
    ' To get to that next row, we have to prestep by the fractional forward distance from A. _Ceil(A.y) - A.y
    prestep_y1 = draw_min_y - A.y

    leg_x1 = A.x + prestep_y1 * legx1_step
    leg_w1 = A.w + prestep_y1 * legw1_step
    leg_u1 = A.u + prestep_y1 * legu1_step
    leg_v1 = A.v + prestep_y1 * legv1_step

    leg_x2 = A.x + prestep_y1 * legx2_step
    leg_w2 = A.w + prestep_y1 * legw2_step
    leg_u2 = A.u + prestep_y1 * legu2_step
    leg_v2 = A.v + prestep_y1 * legv2_step

    ' Inner loop vars
    Static row As Long
    Static col As Long
    Static draw_max_x As Long
    Static tex_z As Single ' 1/w helper (multiply by inverse is faster than dividing each time)
    Static pixel_value As _Unsigned Long ' The ARGB value to write to screen

    ' Stepping along the X direction
    Static delta_x As Single
    Static prestep_x As Single
    Static tex_w_step As Single, tex_u_step As Single, tex_v_step As Single

    ' X Accumulators
    Static tex_w As Single, tex_u As Single, tex_v As Single

    ' Screen Memory Pointers
    Static screen_mem_info As _MEM
    Static screen_next_row_step As _Offset
    Static screen_row_base As _Offset ' Calculated every row
    Static screen_address As _Offset ' Calculated at every starting column
    screen_mem_info = _MemImage(WORK_IMAGE)
    screen_next_row_step = 4 * Size_Render_X

    ' caching of 4 texels in bilinear mode
    Static T1_last_cache As _Unsigned Long
    T1_last_cache = &HFFFFFFFF ' invalidate

    ' Row Loop from top to bottom
    row = draw_min_y
    screen_row_base = screen_mem_info.OFFSET + row * screen_next_row_step
    While row <= draw_max_y

        If row = draw_middle_y Then
            ' Reached Leg 1 knee at B, recalculate Leg 1.
            ' This overwrites Leg 1 to be from B to C. Leg 2 just keeps continuing from A to C.
            delta1.x = C.x - B.x
            delta1.y = C.y - B.y
            delta1.w = C.w - B.w
            delta1.u = C.u - B.u
            delta1.v = C.v - B.v

            If delta1.y = 0.0 Then Exit Sub

            ' Full steps in the y direction from B to C
            legx1_step = delta1.x / delta1.y
            legw1_step = delta1.w / delta1.y
            legu1_step = delta1.u / delta1.y
            legv1_step = delta1.v / delta1.y

            ' 11-4-2022 Prestep Y
            ' Most cases has B lower downscreen than A.
            ' B > A usually. Only one case where B = A.
            prestep_y1 = draw_middle_y - B.y

            ' Re-Initialize DDA start values
            leg_x1 = B.x + prestep_y1 * legx1_step
            leg_w1 = B.w + prestep_y1 * legw1_step
            leg_u1 = B.u + prestep_y1 * legu1_step
            leg_v1 = B.v + prestep_y1 * legv1_step

        End If

        ' Horizontal Scanline
        delta_x = Abs(leg_x2 - leg_x1)
        ' Avoid div/0, this gets tiring.
        If delta_x >= (1 / 2048) Then
            ' Calculate step, start, and end values.
            ' Drawing left to right, as in incrementing from a lower to higher memory address, is usually fastest.
            If leg_x1 < leg_x2 Then
                ' leg 1 is on the left
                tex_w_step = (leg_w2 - leg_w1) / delta_x
                tex_u_step = (leg_u2 - leg_u1) / delta_x
                tex_v_step = (leg_v2 - leg_v1) / delta_x

                ' Set the horizontal starting point to (1)
                col = _Ceil(leg_x1)
                If col < clip_min_x Then col = clip_min_x

                ' Prestep to find pixel starting point
                prestep_x = col - leg_x1
                tex_w = leg_w1 + prestep_x * tex_w_step
                tex_z = 1 / tex_w ' this can be absorbed
                tex_u = leg_u1 + prestep_x * tex_u_step
                tex_v = leg_v1 + prestep_x * tex_v_step

                ' ending point is (2)
                draw_max_x = _Ceil(leg_x2)
                If draw_max_x > clip_max_x Then draw_max_x = clip_max_x

            Else
                ' Things are flipped. leg 1 is on the right.
                tex_w_step = (leg_w1 - leg_w2) / delta_x
                tex_u_step = (leg_u1 - leg_u2) / delta_x
                tex_v_step = (leg_v1 - leg_v2) / delta_x

                ' Set the horizontal starting point to (2)
                col = _Ceil(leg_x2)
                If col < clip_min_x Then col = clip_min_x

                ' Prestep to find pixel starting point
                prestep_x = col - leg_x2
                tex_w = leg_w2 + prestep_x * tex_w_step
                tex_z = 1 / tex_w ' this can be absorbed
                tex_u = leg_u2 + prestep_x * tex_u_step
                tex_v = leg_v2 + prestep_x * tex_v_step

                ' ending point is (1)
                draw_max_x = _Ceil(leg_x1)
                If draw_max_x > clip_max_x Then draw_max_x = clip_max_x

            End If

            ' Draw the Horizontal Scanline
            ' Optimization: before entering this loop, must have done tex_z = 1 / tex_w
            ' Relies on some shared T1 variables over by Texture1
            screen_address = screen_row_base + 4 * col
            While col < draw_max_x

                Static cc As _Unsigned Integer
                Static ccp As _Unsigned Integer
                Static rr As _Unsigned Integer
                Static rrp As _Unsigned Integer

                Static cm5 As Single
                Static rm5 As Single

                ' Recover U and V
                ' Offset so the transition appears in the center of an enlarged texel instead of a corner.
                cm5 = (tex_u * tex_z) - 0.5
                rm5 = (tex_v * tex_z) - 0.5

                ' clamp
                If cm5 < 0.0 Then cm5 = 0.0
                If cm5 >= T1_width_MASK Then
                    ' 15.0 and up
                    cc = T1_width_MASK
                    ccp = T1_width_MASK
                Else
                    ' 0 1 2 .. 13 14.999
                    cc = Int(cm5)
                    ccp = cc + 1
                End If

                ' clamp
                If rm5 < 0.0 Then rm5 = 0.0
                If rm5 >= T1_height_MASK Then
                    ' 15.0 and up
                    rr = T1_height_MASK
                    rrp = T1_height_MASK
                Else
                    rr = Int(rm5)
                    rrp = rr + 1
                End If

                ' 4 point bilinear temp vars
                Static Frac_cc1_FIX7 As Integer
                Static Frac_rr1_FIX7 As Integer
                ' 0 1
                ' . .
                Static bi_r0 As Integer
                Static bi_g0 As Integer
                Static bi_b0 As Integer
                ' . .
                ' 2 3
                Static bi_r1 As Integer
                Static bi_g1 As Integer
                Static bi_b1 As Integer

                Frac_cc1_FIX7 = (cm5 - Int(cm5)) * 128
                Frac_rr1_FIX7 = (rm5 - Int(rm5)) * 128

                ' Caching of 4 texels
                Static T1_this_cache As _Unsigned Long
                Static T1_uv_0_0 As Long
                Static T1_uv_1_0 As Long
                Static T1_uv_0_1 As Long
                Static T1_uv_1_1 As Long

                T1_this_cache = _ShL(rr, 12) Or cc
                If T1_this_cache <> T1_last_cache Then

                    _MemGet T1_mblock, T1_mblock.OFFSET + (cc + rr * T1_width) * 4, T1_uv_0_0
                    _MemGet T1_mblock, T1_mblock.OFFSET + (ccp + rr * T1_width) * 4, T1_uv_1_0
                    _MemGet T1_mblock, T1_mblock.OFFSET + (cc + rrp * T1_width) * 4, T1_uv_0_1
                    _MemGet T1_mblock, T1_mblock.OFFSET + (ccp + rrp * T1_width) * 4, T1_uv_1_1

                    T1_last_cache = T1_this_cache
                End If

                ' determine T1 RGB colors
                bi_r0 = _Red32(T1_uv_0_0)
                bi_r0 = _ShR((_Red32(T1_uv_1_0) - bi_r0) * Frac_cc1_FIX7, 7) + bi_r0

                bi_g0 = _Green32(T1_uv_0_0)
                bi_g0 = _ShR((_Green32(T1_uv_1_0) - bi_g0) * Frac_cc1_FIX7, 7) + bi_g0

                bi_b0 = _Blue32(T1_uv_0_0)
                bi_b0 = _ShR((_Blue32(T1_uv_1_0) - bi_b0) * Frac_cc1_FIX7, 7) + bi_b0

                bi_r1 = _Red32(T1_uv_0_1)
                bi_r1 = _ShR((_Red32(T1_uv_1_1) - bi_r1) * Frac_cc1_FIX7, 7) + bi_r1

                bi_g1 = _Green32(T1_uv_0_1)
                bi_g1 = _ShR((_Green32(T1_uv_1_1) - bi_g1) * Frac_cc1_FIX7, 7) + bi_g1

                bi_b1 = _Blue32(T1_uv_0_1)
                bi_b1 = _ShR((_Blue32(T1_uv_1_1) - bi_b1) * Frac_cc1_FIX7, 7) + bi_b1

                pixel_value = _RGB32(_ShR((bi_r1 - bi_r0) * Frac_rr1_FIX7, 7) + bi_r0, _ShR((bi_g1 - bi_g0) * Frac_rr1_FIX7, 7) + bi_g0, _ShR((bi_b1 - bi_b0) * Frac_rr1_FIX7, 7) + bi_b0)
                _MemPut screen_mem_info, screen_address, pixel_value
                'PSet (col, row), pixel_value

                tex_w = tex_w + tex_w_step
                tex_z = 1 / tex_w ' execution time for this can be absorbed when result not required immediately
                tex_u = tex_u + tex_u_step
                tex_v = tex_v + tex_v_step
                screen_address = screen_address + 4
                col = col + 1
            Wend ' col

        End If ' end div/0 avoidance

        ' DDA next step
        leg_x1 = leg_x1 + legx1_step
        leg_w1 = leg_w1 + legw1_step
        leg_u1 = leg_u1 + legu1_step
        leg_v1 = leg_v1 + legv1_step

        leg_x2 = leg_x2 + legx2_step
        leg_w2 = leg_w2 + legw2_step
        leg_u2 = leg_u2 + legu2_step
        leg_v2 = leg_v2 + legv2_step

        screen_row_base = screen_row_base + screen_next_row_step
        row = row + 1
    Wend ' row

End Sub


Sub ReflectionMapTriangle (A As vertex9, B As vertex9, C As vertex9)
    Static delta2 As vertex9
    Static delta1 As vertex9
    Static draw_min_y As Long, draw_max_y As Long

    ' Sort so that vertex A is on top and C is on bottom.
    ' This seems inverted from math class, but remember that Y increases in value downward on PC monitors
    If B.y < A.y Then
        Swap A, B
    End If
    If C.y < A.y Then
        Swap A, C
    End If
    If C.y < B.y Then
        Swap B, C
    End If

    ' integer window clipping
    draw_min_y = _Ceil(A.y)
    If draw_min_y < clip_min_y Then draw_min_y = clip_min_y
    draw_max_y = _Ceil(C.y) - 1
    If draw_max_y > clip_max_y Then draw_max_y = clip_max_y
    If (draw_max_y - draw_min_y) < 0 Then Exit Sub

    ' Determine the deltas (lengths)
    ' delta 2 is from A to C (the full triangle height)
    delta2.x = C.x - A.x
    delta2.y = C.y - A.y
    delta2.w = C.w - A.w
    delta2.u = C.u - A.u
    delta2.v = C.v - A.v
    delta2.r = C.r - A.r
    delta2.g = C.g - A.g
    delta2.b = C.b - A.b
    delta2.a = C.a - A.a

    ' Avoiding div by 0
    ' Entire Y height less than 1/256 would not have meaningful pixel color change
    If delta2.y < (1 / 256) Then Exit Sub

    ' Determine vertical Y steps for DDA style math
    ' DDA is Digital Differential Analyzer
    ' It is an accumulator that counts from a known start point to an end point, in equal increments defined by the number of steps in-between.
    ' Probably faster nowadays to do the one division at the start, instead of Bresenham, anyway.
    Static legx1_step As Single
    Static legw1_step As Single, legu1_step As Single, legv1_step As Single
    Static legr1_step As Single, legg1_step As Single, legb1_step As Single
    Static lega1_step As Single

    Static legx2_step As Single
    Static legw2_step As Single, legu2_step As Single, legv2_step As Single
    Static legr2_step As Single, legg2_step As Single, legb2_step As Single
    Static lega2_step As Single

    ' Leg 2 steps from A to C (the full triangle height)
    legx2_step = delta2.x / delta2.y
    legw2_step = delta2.w / delta2.y
    legu2_step = delta2.u / delta2.y
    legv2_step = delta2.v / delta2.y
    legr2_step = delta2.r / delta2.y
    legg2_step = delta2.g / delta2.y
    legb2_step = delta2.b / delta2.y
    lega2_step = delta2.a / delta2.y

    ' Leg 1, Draw top to middle
    ' For most triangles, draw downward from the apex A to a knee B.
    ' That knee could be on either the left or right side, but that is handled much later.
    Static draw_middle_y As Long
    draw_middle_y = _Ceil(B.y)
    If draw_middle_y < clip_min_y Then draw_middle_y = clip_min_y
    ' Do not clip B to max_y. Let the y count expire before reaching the knee if it is past bottom of screen.

    ' Leg 1 is from A to B (right now)
    delta1.x = B.x - A.x
    delta1.y = B.y - A.y
    delta1.w = B.w - A.w
    delta1.u = B.u - A.u
    delta1.v = B.v - A.v
    delta1.r = B.r - A.r
    delta1.g = B.g - A.g
    delta1.b = B.b - A.b
    delta1.a = B.a - A.a

    ' If the triangle has no knee, this section gets skipped to avoid divide by 0.
    ' That is okay, because the recalculate Leg 1 from B to C triggers before actually drawing.
    If delta1.y > (1 / 256) Then
        ' Find Leg 1 steps in the y direction from A to B
        legx1_step = delta1.x / delta1.y
        legw1_step = delta1.w / delta1.y
        legu1_step = delta1.u / delta1.y
        legv1_step = delta1.v / delta1.y
        legr1_step = delta1.r / delta1.y
        legg1_step = delta1.g / delta1.y
        legb1_step = delta1.b / delta1.y
        lega1_step = delta1.a / delta1.y
    End If

    ' Y Accumulators
    Static leg_x1 As Single
    Static leg_w1 As Single, leg_u1 As Single, leg_v1 As Single
    Static leg_r1 As Single, leg_g1 As Single, leg_b1 As Single
    Static leg_a1 As Single

    Static leg_x2 As Single
    Static leg_w2 As Single, leg_u2 As Single, leg_v2 As Single
    Static leg_r2 As Single, leg_g2 As Single, leg_b2 As Single
    Static leg_a2 As Single

    ' 11-4-2022 Prestep Y
    Static prestep_y1 As Single
    ' Basically we are sampling pixels on integer exact rows.
    ' But we only are able to know the next row by way of forward interpolation. So always round up.
    ' To get to that next row, we have to prestep by the fractional forward distance from A. _Ceil(A.y) - A.y
    prestep_y1 = draw_min_y - A.y

    leg_x1 = A.x + prestep_y1 * legx1_step
    leg_w1 = A.w + prestep_y1 * legw1_step
    leg_u1 = A.u + prestep_y1 * legu1_step
    leg_v1 = A.v + prestep_y1 * legv1_step
    leg_r1 = A.r + prestep_y1 * legr1_step
    leg_g1 = A.g + prestep_y1 * legg1_step
    leg_b1 = A.b + prestep_y1 * legb1_step
    leg_a1 = A.a + prestep_y1 * lega1_step

    leg_x2 = A.x + prestep_y1 * legx2_step
    leg_w2 = A.w + prestep_y1 * legw2_step
    leg_u2 = A.u + prestep_y1 * legu2_step
    leg_v2 = A.v + prestep_y1 * legv2_step
    leg_r2 = A.r + prestep_y1 * legr2_step
    leg_g2 = A.g + prestep_y1 * legg2_step
    leg_b2 = A.b + prestep_y1 * legb2_step
    leg_a2 = A.a + prestep_y1 * lega2_step

    ' Inner loop vars
    Static row As Long
    Static col As Long
    Static draw_max_x As Long
    Static zbuf_index As _Unsigned Long ' Z-Buffer
    Static tex_z As Single ' 1/w helper (multiply by inverse is faster than dividing each time)
    Static pixel_alpha As Single

    ' Stepping along the X direction
    Static delta_x As Single
    Static prestep_x As Single
    Static tex_w_step As Single, tex_u_step As Single, tex_v_step As Single
    Static tex_r_step As Single, tex_g_step As Single, tex_b_step As Single
    Static tex_a_step As Single

    ' X Accumulators
    Static tex_w As Single, tex_u As Single, tex_v As Single
    Static tex_r As Single, tex_g As Single, tex_b As Single
    Static tex_a As Single

    ' Work Screen Memory Pointers
    Static work_mem_info As _MEM
    Static work_next_row_step As _Offset
    Static work_row_base As _Offset ' Calculated every row
    Static work_address As _Offset ' Calculated at every starting column
    work_mem_info = _MemImage(WORK_IMAGE)
    work_next_row_step = 4 * Size_Render_X

    ' Invalidate texel cache
    T1_last_cache = &HFFFFFFFF

    ' Row Loop from top to bottom
    row = draw_min_y
    work_row_base = work_mem_info.OFFSET + row * work_next_row_step
    While row <= draw_max_y

        If row = draw_middle_y Then
            ' Reached Leg 1 knee at B, recalculate Leg 1.
            ' This overwrites Leg 1 to be from B to C. Leg 2 just keeps continuing from A to C.
            delta1.x = C.x - B.x
            delta1.y = C.y - B.y
            delta1.w = C.w - B.w
            delta1.u = C.u - B.u
            delta1.v = C.v - B.v
            delta1.r = C.r - B.r
            delta1.g = C.g - B.g
            delta1.b = C.b - B.b
            delta1.a = C.a - B.a

            If delta1.y = 0.0 Then Exit Sub

            ' Full steps in the y direction from B to C
            legx1_step = delta1.x / delta1.y
            legw1_step = delta1.w / delta1.y
            legu1_step = delta1.u / delta1.y
            legv1_step = delta1.v / delta1.y
            legr1_step = delta1.r / delta1.y ' vertex color
            legg1_step = delta1.g / delta1.y
            legb1_step = delta1.b / delta1.y
            lega1_step = delta1.a / delta1.y

            ' 11-4-2022 Prestep Y
            ' Most cases has B lower downscreen than A.
            ' B > A usually. Only one case where B = A.
            prestep_y1 = draw_middle_y - B.y

            ' Re-Initialize DDA start values
            leg_x1 = B.x + prestep_y1 * legx1_step
            leg_w1 = B.w + prestep_y1 * legw1_step
            leg_u1 = B.u + prestep_y1 * legu1_step
            leg_v1 = B.v + prestep_y1 * legv1_step
            leg_r1 = B.r + prestep_y1 * legr1_step
            leg_g1 = B.g + prestep_y1 * legg1_step
            leg_b1 = B.b + prestep_y1 * legb1_step
            leg_a1 = B.a + prestep_y1 * lega1_step
        End If

        ' Horizontal Scanline
        delta_x = Abs(leg_x2 - leg_x1)
        ' Avoid div/0, this gets tiring.
        If delta_x >= (1 / 2048) Then
            ' Calculate step, start, and end values.
            ' Drawing left to right, as in incrementing from a lower to higher memory address, is usually fastest.
            If leg_x1 < leg_x2 Then
                ' leg 1 is on the left
                tex_w_step = (leg_w2 - leg_w1) / delta_x
                tex_u_step = (leg_u2 - leg_u1) / delta_x
                tex_v_step = (leg_v2 - leg_v1) / delta_x
                tex_r_step = (leg_r2 - leg_r1) / delta_x
                tex_g_step = (leg_g2 - leg_g1) / delta_x
                tex_b_step = (leg_b2 - leg_b1) / delta_x
                tex_a_step = (leg_a2 - leg_a1) / delta_x

                ' Set the horizontal starting point to (1)
                col = _Ceil(leg_x1)
                If col < clip_min_x Then col = clip_min_x

                ' Prestep to find pixel starting point
                prestep_x = col - leg_x1
                tex_w = leg_w1 + prestep_x * tex_w_step
                tex_z = 1 / tex_w ' this can be absorbed
                tex_u = leg_u1 + prestep_x * tex_u_step
                tex_v = leg_v1 + prestep_x * tex_v_step
                tex_r = leg_r1 + prestep_x * tex_r_step
                tex_g = leg_g1 + prestep_x * tex_g_step
                tex_b = leg_b1 + prestep_x * tex_b_step
                tex_a = leg_a1 + prestep_x * tex_a_step

                ' ending point is (2)
                draw_max_x = _Ceil(leg_x2)
                If draw_max_x > clip_max_x Then draw_max_x = clip_max_x

            Else
                ' Things are flipped. leg 1 is on the right.
                tex_w_step = (leg_w1 - leg_w2) / delta_x
                tex_u_step = (leg_u1 - leg_u2) / delta_x
                tex_v_step = (leg_v1 - leg_v2) / delta_x
                tex_r_step = (leg_r1 - leg_r2) / delta_x
                tex_g_step = (leg_g1 - leg_g2) / delta_x
                tex_b_step = (leg_b1 - leg_b2) / delta_x
                tex_a_step = (leg_a1 - leg_a2) / delta_x

                ' Set the horizontal starting point to (2)
                col = _Ceil(leg_x2)
                If col < clip_min_x Then col = clip_min_x

                ' Prestep to find pixel starting point
                prestep_x = col - leg_x2
                tex_w = leg_w2 + prestep_x * tex_w_step
                tex_z = 1 / tex_w ' this can be absorbed
                tex_u = leg_u2 + prestep_x * tex_u_step
                tex_v = leg_v2 + prestep_x * tex_v_step
                tex_r = leg_r2 + prestep_x * tex_r_step
                tex_g = leg_g2 + prestep_x * tex_g_step
                tex_b = leg_b2 + prestep_x * tex_b_step
                tex_a = leg_a2 + prestep_x * tex_a_step

                ' ending point is (1)
                draw_max_x = _Ceil(leg_x1)
                If draw_max_x > clip_max_x Then draw_max_x = clip_max_x

            End If

            ' Draw the Horizontal Scanline
            ' Optimization: before entering this loop, must have done tex_z = 1 / tex_w
            work_address = work_row_base + 4 * col
            zbuf_index = row * Size_Render_X + col
            While col < draw_max_x

                If Screen_Z_Buffer(zbuf_index) = 0.0 Or tex_z < Screen_Z_Buffer(zbuf_index) Then
                    If (T1_options And T1_option_no_Z_write) = 0 Then
                        Screen_Z_Buffer(zbuf_index) = tex_z + Z_Fight_Bias
                    End If

                    ' YOU HAD THE RIGHT IDEA
                    Static cubemap_index As Integer
                    Static cubemap_u As Single
                    Static cubemap_v As Single
                    ConvertXYZ_to_CubeIUV tex_r * tex_z, tex_g * tex_z, tex_b * tex_z, cubemap_index, cubemap_u, cubemap_v

                    ' Fill in Texture 1 data
                    T1_ImageHandle = SkyBoxRef(cubemap_index)
                    T1_mblock = _MemImage(T1_ImageHandle)
                    T1_width = _Width(T1_ImageHandle): T1_width_MASK = T1_width - 1
                    T1_height = _Height(T1_ImageHandle): T1_height_MASK = T1_height - 1

                    '--- Begin Inline Texel Read
                    ' Originally function ReadTexel3Point& (ccol As Single, rrow As Single)
                    ' Relies on some shared T1 variables over by Texture1
                    Static cc As Integer
                    Static rr As Integer
                    Static cc1 As Integer
                    Static rr1 As Integer

                    Static Frac_cc1 As Single
                    Static Frac_rr1 As Single

                    Static Area_00 As Single
                    Static Area_11 As Single
                    Static Area_2f As Single

                    Static T1_address_pointer As _Offset
                    Static uv_0_0 As _Unsigned Long
                    Static uv_1_1 As _Unsigned Long
                    Static uv_f As _Unsigned Long

                    Static r0 As Integer
                    Static g0 As Integer
                    Static b0 As Integer

                    Static cm5 As Single
                    Static rm5 As Single

                    ' Offset so the transition appears in the center of an enlarged texel instead of a corner.
                    cm5 = (cubemap_u * T1_width) - 0.5
                    rm5 = (cubemap_v * T1_height) - 0.5

                    ' clamp
                    If cm5 < 0.0 Then cm5 = 0.0
                    If cm5 >= T1_width_MASK Then
                        '15.0 and up
                        cc = T1_width_MASK
                        cc1 = T1_width_MASK
                    Else
                        '0 1 2 .. 13 14.999
                        cc = Int(cm5)
                        cc1 = cc + 1
                    End If

                    ' clamp
                    If rm5 < 0.0 Then rm5 = 0.0
                    If rm5 >= T1_height_MASK Then
                        '15.0 and up
                        rr = T1_height_MASK
                        rr1 = T1_height_MASK
                    Else
                        rr = Int(rm5)
                        rr1 = rr + 1
                    End If

                    'uv_0_0 = Texture1(cc, rr)
                    T1_address_pointer = T1_mblock.OFFSET + (cc + rr * T1_width) * 4
                    _MemGet T1_mblock, T1_address_pointer, uv_0_0

                    'uv_1_1 = Texture1(cc1, rr1)
                    T1_address_pointer = T1_mblock.OFFSET + (cc1 + rr1 * T1_width) * 4
                    _MemGet T1_mblock, T1_address_pointer, uv_1_1

                    Frac_cc1 = cm5 - Int(cm5)
                    Frac_rr1 = rm5 - Int(rm5)

                    If Frac_cc1 > Frac_rr1 Then
                        ' top-right
                        ' Area of a triangle = 1/2 * base * height
                        ' Using twice the areas (rectangles) to eliminate a multiply by 1/2 and a later divide by 1/2
                        Area_11 = Frac_rr1
                        Area_00 = 1.0 - Frac_cc1

                        'uv_f = Texture1(cc1, rr)
                        T1_address_pointer = T1_mblock.OFFSET + (cc1 + rr * T1_width) * 4
                        _MemGet T1_mblock, T1_address_pointer, uv_f
                    Else
                        ' bottom-left
                        Area_00 = 1.0 - Frac_rr1
                        Area_11 = Frac_cc1

                        'uv_f = Texture1(cc, rr1)
                        T1_address_pointer = T1_mblock.OFFSET + (cc + rr1 * T1_width) * 4
                        _MemGet T1_mblock, T1_address_pointer, uv_f

                    End If

                    Area_2f = 1.0 - (Area_00 + Area_11) '1.0 here is twice the total triangle area.

                    r0 = _Red32(uv_f) * Area_2f + _Red32(uv_0_0) * Area_00 + _Red32(uv_1_1) * Area_11
                    g0 = _Green32(uv_f) * Area_2f + _Green32(uv_0_0) * Area_00 + _Green32(uv_1_1) * Area_11
                    b0 = _Blue32(uv_f) * Area_2f + _Blue32(uv_0_0) * Area_00 + _Blue32(uv_1_1) * Area_11
                    '--- End Inline Texel Read

                    Static pixel_combine As _Unsigned Long
                    pixel_combine = _RGB32(r0, g0, b0)

                    Static pixel_existing As _Unsigned Long
                    pixel_alpha = tex_a * tex_z
                    If pixel_alpha < 0.998 Then
                        pixel_existing = _MemGet(work_mem_info, work_address, _Unsigned Long)
                        pixel_combine = _RGB32((  _red32(pixel_combine) - _Red32(pixel_existing))   * pixel_alpha + _red32(pixel_existing), _
                                               (_green32(pixel_combine) - _Green32(pixel_existing)) * pixel_alpha + _green32(pixel_existing), _
                                               ( _Blue32(pixel_combine) - _Blue32(pixel_existing))  * pixel_alpha + _blue32(pixel_existing))

                        ' x = (p1 - p0) * ratio + p0 is equivalent to
                        ' x = (1.0 - ratio) * p0 + ratio * p1
                    End If
                    _MemPut work_mem_info, work_address, pixel_combine
                    'If Dither_Selection > 0 Then
                    '    RGB_Dither555 col, row, pixel_value
                    'Else
                    '    PSet (col, row), pixel_value
                    'End If

                End If ' tex_z
                zbuf_index = zbuf_index + 1
                tex_w = tex_w + tex_w_step
                tex_z = 1 / tex_w ' floating point divide can be done in parallel when result not required immediately.
                tex_u = tex_u + tex_u_step
                tex_v = tex_v + tex_v_step
                tex_r = tex_r + tex_r_step
                tex_g = tex_g + tex_g_step
                tex_b = tex_b + tex_b_step
                tex_a = tex_a + tex_a_step
                work_address = work_address + 4
                col = col + 1
            Wend ' col

        End If ' end div/0 avoidance

        ' DDA next step
        leg_x1 = leg_x1 + legx1_step
        leg_w1 = leg_w1 + legw1_step
        leg_u1 = leg_u1 + legu1_step
        leg_v1 = leg_v1 + legv1_step
        leg_r1 = leg_r1 + legr1_step
        leg_g1 = leg_g1 + legg1_step
        leg_b1 = leg_b1 + legb1_step
        leg_a1 = leg_a1 + lega1_step

        leg_x2 = leg_x2 + legx2_step
        leg_w2 = leg_w2 + legw2_step
        leg_u2 = leg_u2 + legu2_step
        leg_v2 = leg_v2 + legv2_step
        leg_r2 = leg_r2 + legr2_step
        leg_g2 = leg_g2 + legg2_step
        leg_b2 = leg_b2 + legb2_step
        leg_a2 = leg_a2 + lega2_step

        work_row_base = work_row_base + work_next_row_step
        row = row + 1
    Wend ' row

End Sub
