Option _Explicit
_Title "Alias Object File 63"
' 2024 Haggarman
'  V63 Selectable Alpha mask edge type.
'  V62 Simplify 3D point perspective projection function. map_d forces on backface.
'  V61 Turntable and other selectable spin rotations.
'  V59 optional -clamp texture coordinates but be aware Windows 10 3D Viewer does not like it.
'  V58 Environment Map and Backface toggle.
'  V56 Near frustum plane clipping. Mainly for room interior models.
'  V53 camera movement speed adjustment using + or - keys.
'  V52 fix mesh being mirrored due to -Z and CCW winding order differences.
'  V49 fix bad last vt on a 4 vertex face, example: f 1/1 2/2 3/3 4/4
'  V44 Load Kd texture maps
'  V43 Pre-rotate the vertexes to gain about 10 ms on large objects.
'  V41 obj illumination model
'  V34 Any size skybox texture dimensions
'  V31 Mirror Reflective Surface
'  V30 Skybox
'  V26 Experiments with using half-angle instead of bounce reflection for specular.
'  V23 Specular gouraud.
'  V19 Re-introduce texture mapping, although you only get red brick for now.
'  V17 is where the directional lighting using vertex normals is functioning.
'
'  Press G to toggle face lighting method, if vertex normals are available.
'  Press E to toggle skybox environment map.
'  Press B to toggle backface rendering.
'  Press S to toggle spin (move the model). J to select motion type. O resets angles.
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
Dim Obj_File_Path As String


' MODIFY THESE if you want.
Obj_File_Path = "" ' "bunny.obj" "cube.obj" "teacup.obj" "spoonfix.obj" empty string shows system open file dialog box.
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

Type mesh_triangle
    i0 As Long ' vertex index
    i1 As Long
    i2 As Long

    vni0 As Long ' vector normal index
    vni1 As Long
    vni2 As Long

    u0 As Single ' texture coords
    v0 As Single
    u1 As Single
    v1 As Single
    u2 As Single
    v2 As Single

    options As _Unsigned Long
    material As Long
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
    illum As Long '  illumination model
    options As _Unsigned Long
    map_Kd As Long ' image handle to diffuse texture map when -2 or lower
    Kd_r As Single ' diffuse color (main color)
    Kd_g As Single
    Kd_b As Single
    Ks_r As Single ' specular color
    Ks_g As Single
    Ks_b As Single
    Ns As Single ' specular power exponent (0 to 1000)
    diaphaneity As Single ' translucency where 1.0 means fully opaque. Strangely, Alias Wavefront documentation calls this dissolve.
    textName As String
End Type

Type texture_catalog_type
    textName As String
    imageHandle As Long
End Type

Const illum_model_constant_color = 0 ' Kd only
Const illum_model_lambertian = 1 ' ambient constant term plus a diffuse shading term for the angle of each light source
Const illum_model_blinn_phong = 2 ' ambient constant term, plus a diffuse and specular shading term for each light source
Const illum_model_reflection_map_raytrace = 3 ' like model 2, plus Ks * (reflection map + raytracing)
Const illum_model_glass_map_raytrace = 4 ' When dissolve < 0.1, specular highlights from lights or reflections become imperceptible.
Const illum_model_fresnel = 5
Const illum_model_refraction = 6 ' uses optical density (Ni) and resulting light is filtered by Tf (transmission filter)
Const illum_model_fresnel_refraction = 7
Const illum_model_reflection_map = 8 ' like model 2, plus Ks * reflection map. (but without raytracing)
Const illum_model_glass_map = 9 ' like model 4, but without raytrace
Const illum_model_shadow_only = 10 ' movie film shadowmatte

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
Frustum_FOV_ratio = 1.0 / Tan(_D2R(Frustum_FOV_deg * 0.5)) ' 90 degrees gives 1.0

Dim matProj(3, 3) As Single
matProj(0, 0) = Frustum_Aspect_Ratio * Frustum_FOV_ratio ' output X = input X * factors. The screen is wider than it is tall.
matProj(1, 1) = -Frustum_FOV_ratio ' output Y = input Y * factors. Negate so that +Y is up
matProj(2, 2) = Frustum_Far / (Frustum_Far - Frustum_Near) ' remap output Z between near and far planes, scale factor applied to input Z.
matProj(3, 2) = (-Frustum_Far * Frustum_Near) / (Frustum_Far - Frustum_Near) ' remap output Z between near and far planes, constant offset.
matProj(2, 3) = 1.0 ' divide outputs X and Y by input Z.
' All other matrix elements assumed 0.0

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
'Fog_color = _RGB32(2, 2, 2)
Fog_R = _Red(Fog_color)
Fog_G = _Green(Fog_color)
Fog_B = _Blue(Fog_color)

' Z Fight has to do with overdrawing on top of the same coplanar surface.
' If it is positive, a newer pixel at the same exact Z will always overdraw the older one.
Dim Shared Z_Fight_Bias
Z_Fight_Bias = -0.000061


' These T1 Texture characteristics are read later on during drawing.
Dim Shared T1_ImageHandle As Long
Dim Shared T1_width As Integer, T1_height As Integer
Dim Shared T1_width_MASK As Integer, T1_height_MASK As Integer
Dim Shared T1_mblock As _MEM
Dim Shared T1_Alpha_Threshold As Integer

' Texture sampling
Dim Shared Texture_options As _Unsigned Long
Const T1_option_clamp_width = 1
Const T1_option_clamp_height = 2
Const T1_option_no_Z_write = 4
Const T1_option_no_backface_cull = 16
Const T1_option_metallic = 32768
Const T1_option_no_T1 = 65536
Const T1_option_alpha_hard_edge = 131072
Const T1_option_alpha_single_bit = 262144
Const oneOver255 = 1.0 / 255.0

' Give sensible defaults to avoid crashes or invisible textures.
' Later optimization in ReadTexel requires these to be powers of 2.
' That means: 2,4,8,16,32,64,128,256...
T1_width = 16: T1_width_MASK = T1_width - 1
T1_height = 16: T1_height_MASK = T1_height - 1
T1_Alpha_Threshold = 250 ' below this alpha channel value, do not update z buffer (0..255)

' Texture1 math
Dim Shared T1_mod_R As Single ' simulate a generic color register
Dim Shared T1_mod_G As Single
Dim Shared T1_mod_B As Single
Dim Shared T1_mod_A As Single
T1_mod_R = 1.0
T1_mod_G = 1.0
T1_mod_B = 1.0
T1_mod_A = 1.0

Dim Shared TextureCatalog(90) As texture_catalog_type
Dim Shared TextureCatalog_nextIndex
TextureCatalog_nextIndex = 0


' Load Skybox
'     +---+
'     | 2 |
' +---+---+---+---+
' | 1 | 4 | 0 | 5 |
' +---+---+---+---+
'     | 3 |
'     +---+

Dim Shared SkyBoxRef(5) As Long
SkyBoxRef(0) = _LoadImage("px.png", 32)
SkyBoxRef(1) = _LoadImage("nx.png", 32)
SkyBoxRef(2) = _LoadImage("py.png", 32)
SkyBoxRef(3) = _LoadImage("ny.png", 32)
SkyBoxRef(4) = _LoadImage("pz.png", 32)
SkyBoxRef(5) = _LoadImage("nz.png", 32)

' Error _LoadImage returns -1 as an invalid handle if it cannot load the image.
Dim refIndex As Integer
For refIndex = 0 To 5
    If SkyBoxRef(refIndex) = -1 Then
        Print "Could not load texture file for skybox face: "; refIndex
        End
    End If
Next refIndex

Dim Shared Sky_Last_Element As Integer
Sky_Last_Element = 11
Dim sky(Sky_Last_Element) As skybox_triangle
Restore SKYBOX
For refIndex = 0 To Sky_Last_Element
    Read sky(refIndex).x0: Read sky(refIndex).y0: Read sky(refIndex).z0
    Read sky(refIndex).x1: Read sky(refIndex).y1: Read sky(refIndex).z1
    Read sky(refIndex).x2: Read sky(refIndex).y2: Read sky(refIndex).z2

    Read sky(refIndex).u0: Read sky(refIndex).v0
    Read sky(refIndex).u1: Read sky(refIndex).v1
    Read sky(refIndex).u2: Read sky(refIndex).v2

    Read sky(refIndex).texture
    ' Fill in Texture 1 data
    T1_ImageHandle = SkyBoxRef(sky(refIndex).texture)
    T1_mblock = _MemImage(T1_ImageHandle)
    T1_width = _Width(T1_ImageHandle): T1_width_MASK = T1_width - 1
    T1_height = _Height(T1_ImageHandle): T1_height_MASK = T1_height - 1

    ' any size
    sky(refIndex).u0 = sky(refIndex).u0 * T1_width: sky(refIndex).v0 = sky(refIndex).v0 * T1_height
    sky(refIndex).u1 = sky(refIndex).u1 * T1_width: sky(refIndex).v1 = sky(refIndex).v1 * T1_height
    sky(refIndex).u2 = sky(refIndex).u2 * T1_width: sky(refIndex).v2 = sky(refIndex).v2 * T1_height
Next refIndex

' Load Mesh
While Obj_File_Path = ""
    Obj_File_Path = _OpenFileDialog$("Load Alias Object File", , "*.OBJ|*.obj")
    If Obj_File_Path = "" Then End
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
Dim Shared Obj_Directory As String
Dim Shared Obj_FileNameOnly As String
Dim Shared Material_File_Name As String
Dim Shared Material_File_Path As String
Dim Shared Materials_Count As Long
Dim Shared Hidden_Materials_Count As Long

' Isolate the object's directory from its full file path.
' This is because an object file refers to other files that need to be loaded.
Dim fpps As Integer
fpps = 0
Dim i As Integer
For i = Len(Obj_File_Path) To 1 Step -1
    If Asc(Obj_File_Path, i) = 92 Or Asc(Obj_File_Path, i) = 47 Then
        ' found a / slash \
        fpps = i
        Exit For
    End If
Next i

If fpps > 0 Then
    Obj_Directory = Left$(Obj_File_Path, fpps)
    Obj_FileNameOnly = Mid$(Obj_File_Path, fpps + 1)
Else
    ' within current working directory
    Obj_Directory = ""
    Obj_FileNameOnly = Obj_File_Path
End If
Print "Loading "; Obj_FileNameOnly

PrescanMesh Obj_File_Path, Mesh_Last_Element, Vertex_Count, TextureCoord_Count, Vtx_Normals_Count, Material_File_Count, Material_File_Name
If Mesh_Last_Element = 0 Then
    Color _RGB(249, 161, 50)
    Print "Error Loading Object File "; Obj_File_Path
    End
End If

Material_File_Path = Obj_Directory + Material_File_Name
If Material_File_Count >= 1 Then
    PrescanMaterialFile Material_File_Path, Materials_Count
    If Materials_Count = 0 Then
        Color _RGB(249, 161, 50)
        Print "Error Loading Material File "; Material_File_Name
        Print "Material Full Path "; Material_File_Path
        End
    End If
End If

' always create at least one material. 0 just creates a single element index 0.
Dim Shared Materials(Materials_Count) As newmtl_type
' sensible defaults to at least show something.
Materials(0).illum = 2
Materials(0).options = T1_option_no_T1
Materials(0).map_Kd = 0 ' invalid handle intentionally
Materials(0).diaphaneity = 1.0
Materials(0).Kd_r = 0.5: Materials(0).Kd_g = 0.5: Materials(0).Kd_b = 0.5
Materials(0).Ks_r = 0.25: Materials(0).Ks_g = 0.25: Materials(0).Ks_b = 0.25
Materials(0).Ns = 18

Dim L As Long
If Material_File_Count >= 1 Then
    LoadMaterialFile Material_File_Path, Materials(), Materials_Count, Hidden_Materials_Count

    If Hidden_Materials_Count > 0 Then
        Dim ratio As Long
        If Hidden_Materials_Count > Materials_Count \ 2 Then ratio = 1 Else ratio = 0
        Dim askhidden$
        askhidden$ = Str$(Hidden_Materials_Count) + " of " + Str$(Materials_Count) + " materials are hidden. Make them visible?"
        If _MessageBox("Material File", askhidden$, "yesno", "question", ratio) = 1 Then
            For L = 1 To Materials_Count
                If Materials(L).diaphaneity = 0.0 Then Materials(L).diaphaneity = 1.0
            Next L
        End If
    End If
Else
    Print "No Material File specified"
End If

' create and start reading in the triangle mesh
Dim Shared mesh_draw_order(Mesh_Last_Element) As Long
Dim Shared mesh(Mesh_Last_Element) As mesh_triangle

' 10-5-2024
Dim VertexList(Vertex_Count) As vec3d
Dim VtxNorms(Vtx_Normals_Count) As vec3d


Dim actor As Integer
Dim tri As Long
tri = 0
For actor = 1 To Actor_Count
    Objects(actor).first = tri + 1
    LoadMesh Obj_File_Path, mesh(), tri, VertexList(), TextureCoord_Count, VtxNorms(), Materials()
    Objects(actor).last = tri
Next actor

' find the furthest out point to be able to pull back camera start position for large objects
Dim index_v As Long
Dim distance As Double
Dim GreatestDistance As Double
GreatestDistance = 0.0
For tri = 1 To Mesh_Last_Element
    index_v = mesh(tri).i0
    distance = VertexList(index_v).x * VertexList(index_v).x + VertexList(index_v).y * VertexList(index_v).y + VertexList(index_v).z * VertexList(index_v).z
    If distance > GreatestDistance Then GreatestDistance = distance

    index_v = mesh(tri).i1
    distance = VertexList(index_v).x * VertexList(index_v).x + VertexList(index_v).y * VertexList(index_v).y + VertexList(index_v).z * VertexList(index_v).z
    If distance > GreatestDistance Then GreatestDistance = distance

    index_v = mesh(tri).i2
    distance = VertexList(index_v).x * VertexList(index_v).x + VertexList(index_v).y * VertexList(index_v).y + VertexList(index_v).z * VertexList(index_v).z
    If distance > GreatestDistance Then GreatestDistance = distance
Next tri
GreatestDistance = Sqr(GreatestDistance)
distance = -2.0 * GreatestDistance
If distance < Camera_Start_Z Then Camera_Start_Z = distance

' Pre-sort the draw order to make solid triangles draw first, then draw those that are transparent or see-thru.
Dim renderPass As Integer
Dim filter_result As Integer
Dim transparencyFactor As Single
L = 1
For renderPass = 0 To 1
    For tri = 1 To Mesh_Last_Element
        transparencyFactor = Materials(mesh(tri).material).diaphaneity
        filter_result = (transparencyFactor < 1.0) Or ((Materials(mesh(tri).material).options And T1_option_no_backface_cull) <> 0)
        If ((renderPass = 0) And (filter_result = 0)) Or ((renderPass = 1) And (filter_result <> 0)) Then
            mesh_draw_order(L) = tri
            L = L + 1
        End If
    Next tri
Next renderPass


' Here are the 3D math and projection vars
Dim object_vertexes(Vertex_Count) As vec3d

' Rotation
Dim matRotObjX(3, 3) As Single
Dim matRotObjY(3, 3) As Single
Dim matRotObjZ(3, 3) As Single
Dim pointRotObj1 As vec3d
Dim pointRotObj2 As vec3d

Dim point0 As vec3d
Dim point1 As vec3d
Dim point2 As vec3d

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

Dim vCameraPsnNext As vec3d
Dim cameraRay0 As vec3d
Dim dotProductCam As Single

' Adjustable camera movement speed
Const CameraSpeedLevel_max = 12
Dim Shared CameraSpeedLookup(CameraSpeedLevel_max) As Single
Dim Shared CameraSpeedLevel As Integer
CameraSpeedLevel = 5
For i = 0 To CameraSpeedLevel_max
    If i And 1 Then
        ' odd fives
        CameraSpeedLookup(i) = 0.5 * powf(10.0, i \ 2 - 2)
    Else
        ' even tens
        CameraSpeedLookup(i) = powf(10.0, i \ 2 - 3)
    End If
Next i

' View Space 2-10-2023
Dim fPitch As Single ' FPS Camera rotation in YZ plane (X)
Dim fYaw As Single ' FPS Camera rotation in XZ plane (Y)
Dim fRoll As Single
Dim Shared matCameraRot(3, 3) As Single

Dim Shared vCameraHomeFwd As vec3d ' Home angle orientation is facing down the Z line.
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
' When developing, set the light source where the camera starts so you don't go insane when trying to get the specular vectors correct.
vLightDir.x = 0.7411916
vLightDir.y = 0.5735765 '+Y is up
vLightDir.z = -0.3487767 ' Camera_Start_Z
Vector3_Normalize vLightDir
Dim Shared Light_Directional As Single
Dim Shared Light_AmbientVal As Single
Light_AmbientVal = 0.24 ' High Ambient just washes out the image but whatever it's part of the light model

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
Dim reflectLightDir0 As vec3d
Dim reflectLightDir1 As vec3d
Dim reflectLightDir2 As vec3d
Dim light_specular_A As Single
Dim light_specular_B As Single
Dim light_specular_C As Single

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
Dim spinAngleDegX As Single
Dim spinAngleDegY As Single
Dim spinAngleDegZ As Single
spinAngleDegX = 0.0
spinAngleDegY = 0.0
spinAngleDegZ = 0.0

' code execution time
Dim start_ms As Double
Dim trimesh_ms As Double
Dim render_ms As Double
Dim render_period_ms As Double

' physics framerate
Dim frametime_fullframe_ms As Double
Dim frametime_fullframethreshold_ms As Double
Dim frametimestamp_now_ms As Double
Dim frametimestamp_prior_ms As Double
Dim frametimestamp_delta_ms As Double
Dim frame_advance As Integer
Dim frame_tracking_size As Integer
frame_tracking_size = 30
Dim frame_ts(frame_tracking_size) As Double
Dim frame_early_polls(frame_tracking_size) As Long

' Main loop stuff
Dim KeyNow As String
Dim ExitCode As Integer
Dim triCount As Integer
Dim Animate_Spin As Integer
Dim Draw_Environment_Map As Integer
Dim Draw_Backface As Integer
Dim Gouraud_Shading_Selection As Integer
Dim Jog_Motion_Selection As Integer
Dim Alpha_Mask_Selection As Integer
Dim thisMaterial As newmtl_type
Dim triloop_input_poll As Long
Dim Shared Pixels_Drawn_This_Frame As Long

main:
$Checking:Off
ExitCode = 0
Animate_Spin = 0
Draw_Environment_Map = -1
Draw_Backface = 0
Gouraud_Shading_Selection = 1
Jog_Motion_Selection = 1
Alpha_Mask_Selection = 0
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
        Select Case Jog_Motion_Selection
            Case 0
                'Print "Zero Orientation"
                spinAngleDegX = 0
                spinAngleDegY = 0
                spinAngleDegZ = 0
                Animate_Spin = 0
            Case 1
                'Print "Turntable Y-Axis"
                spinAngleDegY = spinAngleDegY + frame_advance * 0.2
            Case 2
                'Print "Roll X-Axis"
                spinAngleDegX = spinAngleDegX - frame_advance * 0.25
            Case 3
                'Print "Tumble X and Z"
                spinAngleDegZ = spinAngleDegZ + frame_advance * 0.355
                spinAngleDegX = spinAngleDegX - frame_advance * 0.23
        End Select
    End If
    frame_advance = 0

    ' Set up rotation matrices
    ' _D2R is just a built-in degrees to radians conversion
    Matrix4_MakeRotation_X spinAngleDegX, matRotObjX()
    Matrix4_MakeRotation_Y spinAngleDegY, matRotObjY()
    Matrix4_MakeRotation_Z spinAngleDegZ, matRotObjZ()

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

    _Dest WORK_IMAGE
    _Source WORK_IMAGE

    ' Clear Z-Buffer
    ' This is a qbasic only optimization. it sets the array to zero. it saves 10 ms.
    ReDim Screen_Z_Buffer(Screen_Z_Buffer_MaxElement)

    If Draw_Environment_Map = 0 Then
        Cls , Fog_color
    Else
        ' Draw Skybox 2-28-2023
        For tri = 0 To Sky_Last_Element
            point0.x = sky(tri).x0
            point0.y = sky(tri).y0
            point0.z = sky(tri).z0

            point1.x = sky(tri).x1
            point1.y = sky(tri).y1
            point1.z = sky(tri).z1

            point2.x = sky(tri).x2
            point2.y = sky(tri).y2
            point2.z = sky(tri).z2

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
            vatr0.u = sky(tri).u0: vatr0.v = sky(tri).v0
            vatr1.u = sky(tri).u1: vatr1.v = sky(tri).v1
            vatr2.u = sky(tri).u2: vatr2.v = sky(tri).v2

            ' Clip more often than not in this example
            NearClip pointView0, pointView1, pointView2, pointView3, vatr0, vatr1, vatr2, vatr3, triCount
            If triCount > 0 Then
                ' Fill in Texture 1 data
                T1_ImageHandle = SkyBoxRef(sky(tri).texture)
                T1_mblock = _MemImage(T1_ImageHandle)
                T1_width = _Width(T1_ImageHandle): T1_width_MASK = T1_width - 1
                T1_height = _Height(T1_ImageHandle): T1_height_MASK = T1_height - 1

                ' Project triangles from 3D -----------------> 2D
                ProjectPerspectiveVector4 pointView0, matProj(), pointProj0
                ProjectPerspectiveVector4 pointView1, matProj(), pointProj1
                ProjectPerspectiveVector4 pointView2, matProj(), pointProj2

                ' Slide to center, then Scale into viewport
                SX0 = (pointProj0.x + 1) * halfWidth
                SY0 = (pointProj0.y + 1) * halfHeight

                SX2 = (pointProj2.x + 1) * halfWidth
                SY2 = (pointProj2.y + 1) * halfHeight

                ' Early scissor reject
                If pointProj0.x > 1.0 And pointProj1.x > 1.0 And pointProj2.x > 1.0 Then GoTo Lbl_SkipEnv012
                If pointProj0.x < -1.0 And pointProj1.x < -1.0 And pointProj2.x < -1.0 Then GoTo Lbl_SkipEnv012
                If pointProj0.y > 1.0 And pointProj1.y > 1.0 And pointProj2.y > 1.0 Then GoTo Lbl_SkipEnv012
                If pointProj0.y < -1.0 And pointProj1.y < -1.0 And pointProj2.y < -1.0 Then GoTo Lbl_SkipEnv012

                SX1 = (pointProj1.x + 1) * halfWidth
                SY1 = (pointProj1.y + 1) * halfHeight

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

                TexturedNonlitTriangle vertexA, vertexB, vertexC
            End If

            Lbl_SkipEnv012:
            If triCount = 2 Then

                ProjectPerspectiveVector4 pointView3, matProj(), pointProj3

                ' Late scissor reject
                If (pointProj0.x > 1.0) And (pointProj2.x > 1.0) And (pointProj3.x > 1.0) Then GoTo Lbl_SkipEnv023
                If (pointProj0.x < -1.0) And (pointProj2.x < -1.0) And (pointProj3.x < -1.0) Then GoTo Lbl_SkipEnv023
                If (pointProj0.y > 1.0) And (pointProj2.y > 1.0) And (pointProj3.y > 1.0) Then GoTo Lbl_SkipEnv023
                If (pointProj0.y < -1.0) And (pointProj2.y < -1.0) And (pointProj3.y < -1.0) Then GoTo Lbl_SkipEnv023

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
            End If
            Lbl_SkipEnv023:
        Next tri
    End If ' skybox

    ' pre-rotate the vertexes
    ' object_vertexes() = T.F. vertexList()
    For tri = 1 To Vertex_Count
        ' Rotate in Z-Axis
        Multiply_Vector3_Matrix4 VertexList(tri), matRotObjZ(), pointRotObj1
        ' Rotate in X-Axis
        Multiply_Vector3_Matrix4 pointRotObj1, matRotObjX(), pointRotObj2
        ' Rotate in Y-Axis
        Multiply_Vector3_Matrix4 pointRotObj2, matRotObjY(), object_vertexes(tri)
    Next tri

    If Vtx_Normals_Count > 0 Then
        ' Vertex normals can be rotated around their origin and still retain their effectiveness.
        ' object_vtx_normals() = T.F. VtxNorms()
        For tri = 1 To Vtx_Normals_Count
            ' Rotate in Z-Axis
            Multiply_Vector3_Matrix4 VtxNorms(tri), matRotObjZ(), pointRotObj1
            ' Rotate in X-Axis
            Multiply_Vector3_Matrix4 pointRotObj1, matRotObjX(), pointRotObj2
            ' Rotate in Y-Axis
            Multiply_Vector3_Matrix4 pointRotObj2, matRotObjY(), object_vtx_normals(tri)
        Next tri
    End If

    triloop_input_poll = 25000 ' because of all of the other stuff above, trigger the first check early.
    vCameraPsnNext = vCameraPsn ' capture where the camera is before starting triangle draw loop
    trimesh_ms = Timer(0.001)
    Pixels_Drawn_This_Frame = 0
    ReDim frame_early_polls(frame_tracking_size) As Long

    ' Draw Triangles
    For L = Objects(actor).first To Objects(actor).last
        tri = mesh_draw_order(L)

        ' material
        thisMaterial = Materials(mesh(tri).material)
        ' Combine the material options with this specific triangle's options.
        ' Reason is that a given tri face could possibly omit texture coordinates or backface could be forced on.
        Select Case Alpha_Mask_Selection
            Case 0
                ' (A)lpha Soft Edge
                Texture_options = thisMaterial.options Or mesh(tri).options
            Case 1
                ' (A)lpha Hard Edge
                Texture_options = thisMaterial.options Or mesh(tri).options Or T1_option_alpha_hard_edge
            Case 2
                ' (A)lpha 1-bit Edge
                Texture_options = thisMaterial.options Or mesh(tri).options Or T1_option_alpha_single_bit
        End Select

        ' the object vertexes have been rotated already, so we just need to reference them
        pointWorld0 = object_vertexes(mesh(tri).i0)
        pointWorld1 = object_vertexes(mesh(tri).i1)
        pointWorld2 = object_vertexes(mesh(tri).i2)

        ' Part 2 (Triangle Surface Normal Calculation)
        CalcSurfaceNormal_3Point pointWorld0, pointWorld1, pointWorld2, tri_normal
        Vector3_Delta vCameraPsn, pointWorld0, cameraRay0 ' be careful, this is not normalized.
        dotProductCam = Vector3_DotProduct!(tri_normal, cameraRay0) ' only interested here in the sign

        If Draw_Backface _Orelse (dotProductCam < 0.0) _Orelse (Texture_options And T1_option_no_backface_cull) Then
            ' Front facing is negative dot product because obj vertex is specified with counter-clockwise (CCW) winding.
            ' Convert World Space --> View Space
            Multiply_Vector3_Matrix4 pointWorld0, matView(), pointView0
            Multiply_Vector3_Matrix4 pointWorld1, matView(), pointView1
            Multiply_Vector3_Matrix4 pointWorld2, matView(), pointView2

            ' Skip if all Z is too close
            If (pointView0.z < Frustum_Near) _Andalso (pointView1.z < Frustum_Near) _Andalso (pointView2.z < Frustum_Near) Then GoTo Lbl_SkipTriAll

            ' Texture 1
            T1_mod_A = thisMaterial.diaphaneity ' all illumination models use d

            If (Texture_options And T1_option_no_T1) = 0 Then
                ' still need to check if texture loaded
                If thisMaterial.map_Kd >= -1 Then
                    ' bad texture handle, just use diffuse color
                    Texture_options = Texture_options Or T1_option_no_T1
                Else
                    ' Fill in Texture 1 data
                    T1_ImageHandle = thisMaterial.map_Kd
                    T1_mblock = _MemImage(T1_ImageHandle)
                    T1_width = _Width(T1_ImageHandle): T1_width_MASK = T1_width - 1
                    T1_height = _Height(T1_ImageHandle): T1_height_MASK = T1_height - 1

                    ' obj vt (0, 0) is bottom left of a texture, but our texture origin is top left.
                    vatr0.u = mesh(tri).u0 * T1_width
                    vatr0.v = T1_height - mesh(tri).v0 * T1_height
                    vatr1.u = mesh(tri).u1 * T1_width
                    vatr1.v = T1_height - mesh(tri).v1 * T1_height
                    vatr2.u = mesh(tri).u2 * T1_width
                    vatr2.v = T1_height - mesh(tri).v2 * T1_height
                End If
            End If

            ' Lighting
            Select Case thisMaterial.illum
                Case illum_model_constant_color
                    ' Kd only

                    If Texture_options And T1_option_no_T1 Then
                        ' define as 8 bit values
                        face_light_r = 255.0 * thisMaterial.Kd_r
                        face_light_g = 255.0 * thisMaterial.Kd_g
                        face_light_b = 255.0 * thisMaterial.Kd_b

                        vatr0.r = face_light_r
                        vatr0.g = face_light_g
                        vatr0.b = face_light_b

                        vatr1.r = face_light_r
                        vatr1.g = face_light_g
                        vatr1.b = face_light_b

                        vatr2.r = face_light_r
                        vatr2.g = face_light_g
                        vatr2.b = face_light_b
                    Else
                        ' draw diffuse texture as-is
                        vatr0.r = 1.0
                        vatr0.g = 1.0
                        vatr0.b = 1.0

                        vatr1.r = 1.0
                        vatr1.g = 1.0
                        vatr1.b = 1.0

                        vatr2.r = 1.0
                        vatr2.g = 1.0
                        vatr2.b = 1.0
                    End If


                Case illum_model_lambertian
                    ' ambient constant term plus a diffuse shading term for the angle of each light source

                    Select Case Gouraud_Shading_Selection
                        Case 0:
                            ' Flat face shading with runtime normal calculation
                            Light_Directional = -Vector3_DotProduct!(tri_normal, vLightDir) ' CCW winding
                            If Light_Directional < 0.0 Then Light_Directional = 0.0

                            If Texture_options And T1_option_no_T1 Then
                                ' define as 8 bit values
                                face_light_r = 255.0 * thisMaterial.Kd_r * (Light_Directional + Light_AmbientVal)
                                face_light_g = 255.0 * thisMaterial.Kd_g * (Light_Directional + Light_AmbientVal)
                                face_light_b = 255.0 * thisMaterial.Kd_b * (Light_Directional + Light_AmbientVal)
                            Else
                                ' range from 0 to 1
                                face_light_r = thisMaterial.Kd_r * (Light_Directional + Light_AmbientVal)
                                face_light_g = thisMaterial.Kd_g * (Light_Directional + Light_AmbientVal)
                                face_light_b = thisMaterial.Kd_b * (Light_Directional + Light_AmbientVal)
                            End If

                            vatr0.r = face_light_r
                            vatr0.g = face_light_g
                            vatr0.b = face_light_b

                            vatr1.r = face_light_r
                            vatr1.g = face_light_g
                            vatr1.b = face_light_b

                            vatr2.r = face_light_r
                            vatr2.g = face_light_g
                            vatr2.b = face_light_b

                        Case 1:
                            ' Smooth shading with precalculated normals

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

                            If Texture_options And T1_option_no_T1 Then
                                ' define as 8 bit values
                                vatr0.r = 255.0 * thisMaterial.Kd_r * (light_directional_A + Light_AmbientVal)
                                vatr0.g = 255.0 * thisMaterial.Kd_g * (light_directional_A + Light_AmbientVal)
                                vatr0.b = 255.0 * thisMaterial.Kd_b * (light_directional_A + Light_AmbientVal)

                                vatr1.r = 255.0 * thisMaterial.Kd_r * (light_directional_B + Light_AmbientVal)
                                vatr1.g = 255.0 * thisMaterial.Kd_g * (light_directional_B + Light_AmbientVal)
                                vatr1.b = 255.0 * thisMaterial.Kd_b * (light_directional_B + Light_AmbientVal)

                                vatr2.r = 255.0 * thisMaterial.Kd_r * (light_directional_C + Light_AmbientVal)
                                vatr2.g = 255.0 * thisMaterial.Kd_g * (light_directional_C + Light_AmbientVal)
                                vatr2.b = 255.0 * thisMaterial.Kd_b * (light_directional_C + Light_AmbientVal)
                            Else
                                ' range from 0 to 1
                                vatr0.r = thisMaterial.Kd_r * (light_directional_A + Light_AmbientVal)
                                vatr0.g = thisMaterial.Kd_g * (light_directional_A + Light_AmbientVal)
                                vatr0.b = thisMaterial.Kd_b * (light_directional_A + Light_AmbientVal)

                                vatr1.r = thisMaterial.Kd_r * (light_directional_B + Light_AmbientVal)
                                vatr1.g = thisMaterial.Kd_g * (light_directional_B + Light_AmbientVal)
                                vatr1.b = thisMaterial.Kd_b * (light_directional_B + Light_AmbientVal)

                                vatr2.r = thisMaterial.Kd_r * (light_directional_C + Light_AmbientVal)
                                vatr2.g = thisMaterial.Kd_g * (light_directional_C + Light_AmbientVal)
                                vatr2.b = thisMaterial.Kd_b * (light_directional_C + Light_AmbientVal)
                            End If
                    End Select


                Case illum_model_blinn_phong
                    ' ambient constant term, plus a diffuse and specular shading term for each light source
                    Select Case Gouraud_Shading_Selection
                        Case 0:
                            ' Flat face shading with runtime normal calculation

                            ' Directional light 1-17-2023
                            Light_Directional = -Vector3_DotProduct!(tri_normal, vLightDir) ' CCW winding
                            If Light_Directional < 0.0 Then Light_Directional = 0.0

                            ' Specular light
                            ' Instead of a normalized light position pointing out from origin, needs to be pointing inward towards the reflective surface.
                            vLightSpec.x = -vLightDir.x: vLightSpec.y = -vLightDir.y: vLightSpec.z = -vLightDir.z
                            Vector3_Reflect vLightSpec, tri_normal, reflectLightDir0
                            'Vector3_Normalize reflectLightDir0

                            ' cameraRay0 was already calculated for backface removal.
                            Vector3_Normalize cameraRay0
                            light_specular_A = Vector3_DotProduct!(reflectLightDir0, cameraRay0)
                            If light_specular_A > 0.0 Then
                                ' this power thing only works because it should range from 0..1 again.
                                ' so what it actually does is a higher power pushes the number towards 0 and makes the rolloff steeper.
                                light_specular_A = powf(light_specular_A, thisMaterial.Ns)
                            Else
                                light_specular_A = 0.0
                            End If

                            If Texture_options And T1_option_no_T1 Then
                                ' define as 8 bit values
                                face_light_r = 255.0 * (thisMaterial.Kd_r * (Light_Directional + Light_AmbientVal) + thisMaterial.Ks_r * light_specular_A)
                                face_light_g = 255.0 * (thisMaterial.Kd_g * (Light_Directional + Light_AmbientVal) + thisMaterial.Ks_g * light_specular_A)
                                face_light_b = 255.0 * (thisMaterial.Kd_b * (Light_Directional + Light_AmbientVal) + thisMaterial.Ks_b * light_specular_A)
                            Else
                                face_light_r = thisMaterial.Kd_r * (Light_Directional + Light_AmbientVal) + thisMaterial.Ks_r * light_specular_A
                                face_light_g = thisMaterial.Kd_g * (Light_Directional + Light_AmbientVal) + thisMaterial.Ks_g * light_specular_A
                                face_light_b = thisMaterial.Kd_b * (Light_Directional + Light_AmbientVal) + thisMaterial.Ks_b * light_specular_A
                            End If

                            vatr0.r = face_light_r
                            vatr0.g = face_light_g
                            vatr0.b = face_light_b

                            vatr1.r = face_light_r
                            vatr1.g = face_light_g
                            vatr1.b = face_light_b

                            vatr2.r = face_light_r
                            vatr2.g = face_light_g
                            vatr2.b = face_light_b

                        Case 1:
                            ' Smooth shading with precalculated normals

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

                            Vector3_Reflect vLightSpec, vertex_normal_A, reflectLightDir0
                            Vector3_Reflect vLightSpec, vertex_normal_B, reflectLightDir1
                            Vector3_Reflect vLightSpec, vertex_normal_C, reflectLightDir2

                            ' A
                            ' cameraRay0 was already calculated for backface removal.
                            Vector3_Normalize cameraRay0
                            light_specular_A = Vector3_DotProduct!(reflectLightDir0, cameraRay0)
                            If light_specular_A > 0.0 Then
                                ' this power thing only works because it should range from 0..1 again.
                                ' so what it actually does is a higher power pushes the number towards 0 and makes the rolloff steeper.
                                light_specular_A = powf(light_specular_A, thisMaterial.Ns)
                            Else
                                light_specular_A = 0.0
                            End If

                            ' B
                            Vector3_Delta vCameraPsn, pointWorld1, cameraRay1
                            Vector3_Normalize cameraRay1
                            light_specular_B = Vector3_DotProduct!(reflectLightDir1, cameraRay1)
                            If light_specular_B > 0.0 Then
                                light_specular_B = powf(light_specular_B, thisMaterial.Ns)
                            Else
                                light_specular_B = 0.0
                            End If

                            ' C
                            Vector3_Delta vCameraPsn, pointWorld2, cameraRay2
                            Vector3_Normalize cameraRay2
                            light_specular_C = Vector3_DotProduct!(reflectLightDir2, cameraRay2)
                            If light_specular_C > 0.0 Then
                                light_specular_C = powf(light_specular_C, thisMaterial.Ns)
                            Else
                                light_specular_C = 0.0
                            End If

                            If Texture_options And T1_option_no_T1 Then
                                ' define as 8 bit values
                                vatr0.r = 255.0 * (thisMaterial.Kd_r * (light_directional_A + Light_AmbientVal) + thisMaterial.Ks_r * light_specular_A)
                                vatr0.g = 255.0 * (thisMaterial.Kd_g * (light_directional_A + Light_AmbientVal) + thisMaterial.Ks_g * light_specular_A)
                                vatr0.b = 255.0 * (thisMaterial.Kd_b * (light_directional_A + Light_AmbientVal) + thisMaterial.Ks_b * light_specular_A)

                                vatr1.r = 255.0 * (thisMaterial.Kd_r * (light_directional_B + Light_AmbientVal) + thisMaterial.Ks_r * light_specular_B)
                                vatr1.g = 255.0 * (thisMaterial.Kd_g * (light_directional_B + Light_AmbientVal) + thisMaterial.Ks_g * light_specular_B)
                                vatr1.b = 255.0 * (thisMaterial.Kd_b * (light_directional_B + Light_AmbientVal) + thisMaterial.Ks_b * light_specular_B)

                                vatr2.r = 255.0 * (thisMaterial.Kd_r * (light_directional_C + Light_AmbientVal) + thisMaterial.Ks_r * light_specular_C)
                                vatr2.g = 255.0 * (thisMaterial.Kd_g * (light_directional_C + Light_AmbientVal) + thisMaterial.Ks_g * light_specular_C)
                                vatr2.b = 255.0 * (thisMaterial.Kd_b * (light_directional_C + Light_AmbientVal) + thisMaterial.Ks_b * light_specular_C)
                            Else
                                ' range from 0 to 1
                                vatr0.r = thisMaterial.Kd_r * (light_directional_A + Light_AmbientVal) + thisMaterial.Ks_r * light_specular_A
                                vatr0.g = thisMaterial.Kd_g * (light_directional_A + Light_AmbientVal) + thisMaterial.Ks_g * light_specular_A
                                vatr0.b = thisMaterial.Kd_b * (light_directional_A + Light_AmbientVal) + thisMaterial.Ks_b * light_specular_A

                                vatr1.r = thisMaterial.Kd_r * (light_directional_B + Light_AmbientVal) + thisMaterial.Ks_r * light_specular_B
                                vatr1.g = thisMaterial.Kd_g * (light_directional_B + Light_AmbientVal) + thisMaterial.Ks_g * light_specular_B
                                vatr1.b = thisMaterial.Kd_b * (light_directional_B + Light_AmbientVal) + thisMaterial.Ks_b * light_specular_B

                                vatr2.r = thisMaterial.Kd_r * (light_directional_C + Light_AmbientVal) + thisMaterial.Ks_r * light_specular_C
                                vatr2.g = thisMaterial.Kd_g * (light_directional_C + Light_AmbientVal) + thisMaterial.Ks_g * light_specular_C
                                vatr2.b = thisMaterial.Kd_b * (light_directional_C + Light_AmbientVal) + thisMaterial.Ks_b * light_specular_C
                            End If
                    End Select


                Case illum_model_reflection_map, illum_model_glass_map
                    ' Mirror
                    ' the reflection map is modulated by the specular factor
                    T1_mod_R = thisMaterial.Ks_r
                    T1_mod_G = thisMaterial.Ks_g
                    T1_mod_B = thisMaterial.Ks_b

                    Select Case Gouraud_Shading_Selection
                        Case 0:
                            ' Flat face shading
                            vertex_normal_A = tri_normal
                            vertex_normal_B = tri_normal
                            vertex_normal_C = tri_normal
                        Case 1:
                            ' 6-15-2024 pre-rotated normals
                            vertex_normal_A = object_vtx_normals(mesh(tri).vni0)
                            vertex_normal_B = object_vtx_normals(mesh(tri).vni1)
                            vertex_normal_C = object_vtx_normals(mesh(tri).vni2)
                    End Select

                    Vector3_Delta pointWorld0, vCameraPsn, cameraRay0
                    Vector3_Reflect cameraRay0, vertex_normal_A, envMapReflectionRayA

                    Vector3_Delta pointWorld1, vCameraPsn, cameraRay1
                    Vector3_Reflect cameraRay1, vertex_normal_B, envMapReflectionRayB

                    Vector3_Delta pointWorld2, vCameraPsn, cameraRay2
                    Vector3_Reflect cameraRay2, vertex_normal_C, envMapReflectionRayC

                    vatr0.r = envMapReflectionRayA.x ' RGB holds (X,Y,Z) instead
                    vatr0.g = envMapReflectionRayA.y
                    vatr0.b = envMapReflectionRayA.z

                    vatr1.r = envMapReflectionRayB.x
                    vatr1.g = envMapReflectionRayB.y
                    vatr1.b = envMapReflectionRayB.z

                    vatr2.r = envMapReflectionRayC.x
                    vatr2.g = envMapReflectionRayC.y
                    vatr2.b = envMapReflectionRayC.z

                    Texture_options = Texture_options Or T1_option_metallic
                Case Else:
            End Select

            ' Clip if any Z is too close. Assumption is that near clip is uncommon.
            ' If there is a lot of near clipping going on, please remove this precheck and just always call NearClip.
            If (pointView0.z < Frustum_Near) _Orelse (pointView1.z < Frustum_Near) _Orelse (pointView2.z < Frustum_Near) Then
                NearClip pointView0, pointView1, pointView2, pointView3, vatr0, vatr1, vatr2, vatr3, triCount
                If triCount = 0 Then GoTo Lbl_SkipTriAll
            Else
                triCount = 1
            End If

            ' Project triangles from 3D -----------------> 2D
            ProjectPerspectiveVector4 pointView0, matProj(), pointProj0
            ProjectPerspectiveVector4 pointView1, matProj(), pointProj1
            ProjectPerspectiveVector4 pointView2, matProj(), pointProj2

            ' Slide to center, then Scale into viewport
            SX0 = (pointProj0.x + 1) * halfWidth
            SY0 = (pointProj0.y + 1) * halfHeight
            SX2 = (pointProj2.x + 1) * halfWidth
            SY2 = (pointProj2.y + 1) * halfHeight

            ' Early scissor reject
            If pointProj0.x > 1.0 And pointProj1.x > 1.0 And pointProj2.x > 1.0 Then GoTo Lbl_Skip012
            If pointProj0.x < -1.0 And pointProj1.x < -1.0 And pointProj2.x < -1.0 Then GoTo Lbl_Skip012
            If pointProj0.y > 1.0 And pointProj1.y > 1.0 And pointProj2.y > 1.0 Then GoTo Lbl_Skip012
            If pointProj0.y < -1.0 And pointProj1.y < -1.0 And pointProj2.y < -1.0 Then GoTo Lbl_Skip012

            ' This is unique to triangle 012
            SX1 = (pointProj1.x + 1) * halfWidth
            SY1 = (pointProj1.y + 1) * halfHeight

            ' Load Vertex Lists
            vertexA.x = SX0
            vertexA.y = SY0
            vertexA.w = pointProj0.w ' depth

            vertexB.x = SX1
            vertexB.y = SY1
            vertexB.w = pointProj1.w ' depth

            vertexC.x = SX2
            vertexC.y = SY2
            vertexC.w = pointProj2.w ' depth

            If Texture_options And T1_option_metallic Then
                vertexA.r = vatr0.r * pointProj0.w
                vertexA.g = vatr0.g * pointProj0.w
                vertexA.b = vatr0.b * pointProj0.w

                vertexB.r = vatr1.r * pointProj1.w
                vertexB.g = vatr1.g * pointProj1.w
                vertexB.b = vatr1.b * pointProj1.w

                vertexC.r = vatr2.r * pointProj2.w
                vertexC.g = vatr2.g * pointProj2.w
                vertexC.b = vatr2.b * pointProj2.w
                ReflectionMapTriangle vertexA, vertexB, vertexC
            Else
                vertexA.r = vatr0.r
                vertexA.g = vatr0.g
                vertexA.b = vatr0.b
                vertexB.r = vatr1.r
                vertexB.g = vatr1.g
                vertexB.b = vatr1.b
                vertexC.r = vatr2.r
                vertexC.g = vatr2.g
                vertexC.b = vatr2.b
                If Texture_options And T1_option_no_T1 Then
                    VertexColorAlphaTriangle vertexA, vertexB, vertexC
                Else
                    vertexA.u = vatr0.u * pointProj0.w
                    vertexA.v = vatr0.v * pointProj0.w
                    vertexB.u = vatr1.u * pointProj1.w
                    vertexB.v = vatr1.v * pointProj1.w
                    vertexC.u = vatr2.u * pointProj2.w
                    vertexC.v = vatr2.v * pointProj2.w
                    TextureWithAlphaTriangle vertexA, vertexB, vertexC
                End If
            End If ' metallic


            Lbl_Skip012:
            If triCount = 2 Then
                ProjectPerspectiveVector4 pointView3, matProj(), pointProj3

                ' Late scissor reject
                If (pointProj0.x > 1.0) And (pointProj2.x > 1.0) And (pointProj3.x > 1.0) Then GoTo Lbl_SkipTriAll
                If (pointProj0.x < -1.0) And (pointProj2.x < -1.0) And (pointProj3.x < -1.0) Then GoTo Lbl_SkipTriAll
                If (pointProj0.y > 1.0) And (pointProj2.y > 1.0) And (pointProj3.y > 1.0) Then GoTo Lbl_SkipTriAll
                If (pointProj0.y < -1.0) And (pointProj2.y < -1.0) And (pointProj3.y < -1.0) Then GoTo Lbl_SkipTriAll

                ' This is unique to triangle 023
                SX3 = (pointProj3.x + 1) * halfWidth
                SY3 = (pointProj3.y + 1) * halfHeight

                ' Reload Vertex Lists
                vertexA.x = SX0
                vertexA.y = SY0
                vertexA.w = pointProj0.w ' depth

                vertexB.x = SX2
                vertexB.y = SY2
                vertexB.w = pointProj2.w ' depth

                vertexC.x = SX3
                vertexC.y = SY3
                vertexC.w = pointProj3.w ' depth

                If Texture_options And T1_option_metallic Then
                    vertexA.r = vatr0.r * pointProj0.w
                    vertexA.g = vatr0.g * pointProj0.w
                    vertexA.b = vatr0.b * pointProj0.w

                    vertexB.r = vatr2.r * pointProj2.w
                    vertexB.g = vatr2.g * pointProj2.w
                    vertexB.b = vatr2.b * pointProj2.w

                    vertexC.r = vatr3.r * pointProj3.w
                    vertexC.g = vatr3.g * pointProj3.w
                    vertexC.b = vatr3.b * pointProj3.w
                    ReflectionMapTriangle vertexA, vertexB, vertexC
                Else
                    vertexA.r = vatr0.r
                    vertexA.g = vatr0.g
                    vertexA.b = vatr0.b
                    vertexB.r = vatr2.r
                    vertexB.g = vatr2.g
                    vertexB.b = vatr2.b
                    vertexC.r = vatr3.r
                    vertexC.g = vatr3.g
                    vertexC.b = vatr3.b
                    If Texture_options And T1_option_no_T1 Then
                        VertexColorAlphaTriangle vertexA, vertexB, vertexC
                    Else
                        vertexA.u = vatr0.u * pointProj0.w
                        vertexA.v = vatr0.v * pointProj0.w
                        vertexB.u = vatr2.u * pointProj2.w
                        vertexB.v = vatr2.v * pointProj2.w
                        vertexC.u = vatr3.u * pointProj3.w
                        vertexC.v = vatr3.v * pointProj3.w
                        TextureWithAlphaTriangle vertexA, vertexB, vertexC
                    End If
                End If ' metallic
            End If ' tricount=2

        End If ' visible according to dotProductCam

        ' Improve camera control feeling by polling the keyboard every so often during this main triangle draw loop.
        triloop_input_poll = triloop_input_poll + 1
        If triloop_input_poll < 30000 GoTo Lbl_SkipTriAll
        triloop_input_poll = triloop_input_poll - 6000 ' in case too early
        frame_early_polls(frame_advance) = frame_early_polls(frame_advance) + 1

        frametimestamp_now_ms = Timer(0.001) ' this function is slow
        If frametimestamp_now_ms - frametimestamp_prior_ms < 0.0 Then
            ' timer rollover
            ' without over-analyzing just use the previous delta, even if it is somewhat wrong it is a better guess than 0.
            frametimestamp_prior_ms = frametimestamp_now_ms - frametimestamp_delta_ms
        Else
            frametimestamp_delta_ms = frametimestamp_now_ms - frametimestamp_prior_ms
        End If
        While frametimestamp_delta_ms > frametime_fullframethreshold_ms
            frame_advance = frame_advance + 1
            If frame_advance > frame_tracking_size Then frame_advance = frame_tracking_size
            frame_ts(frame_advance) = frametimestamp_delta_ms

            frametimestamp_prior_ms = frametimestamp_prior_ms + frametime_fullframe_ms
            frametimestamp_delta_ms = frametimestamp_delta_ms - frametime_fullframe_ms

            CameraPoll vCameraPsnNext, fYaw, fPitch
            triloop_input_poll = 0 ' did the poll so go back
        Wend ' frametime

        Lbl_SkipTriAll:
    Next L

    render_ms = Timer(.001)

    _PutImage , WORK_IMAGE, DISP_IMAGE
    _Dest DISP_IMAGE
    Locate 1, 1
    Color _RGB32(177, 227, 255)
    Print Using "render time #.###"; render_ms - start_ms
    Color _RGB32(249, 244, 17)
    Print "ESC to exit. ";
    Color _RGB32(233)
    Print "Arrow Keys Move. -Speed+:"; CameraSpeedLookup(CameraSpeedLevel)

    If Jog_Motion_Selection <> 0 Then
        If Animate_Spin Then
            Print "(S)top Spin  ";
        Else
            Print "(S)tart Spin ";
        End If
    End If

    Print "(J)og Type: ";
    Select Case Jog_Motion_Selection
        Case 0
            Print "Zero Orientation"
        Case 1
            Print "Turntable Y-Axis"
        Case 2
            Print "Roll X-Axis"
        Case 3
            Print "Tumble X and Z"
        Case Else
            Print Jog_Motion_Selection
    End Select

    If Draw_Environment_Map Then
        Print "(E)nvironment Map on ";
    Else
        Print "(E)nvironment Map off";
    End If

    If Draw_Backface Then
        Print " (B)ackface on ";
    Else
        Print " (B)ackface off";
    End If

    Select Case Alpha_Mask_Selection
        Case 0
            Print " (A)lpha Soft Edge"
        Case 1
            Print " (A)lpha Hard Edge"
        Case 2
            Print " (A)lpha 1-bit Edge"
    End Select

    Print "Press G for Lighting: ";
    Select Case Gouraud_Shading_Selection
        Case 0
            Print "Flat using face normals"
        Case 1
            Print "Gouraud using vertex normals"
        Case 2
            Print "Mirror"
    End Select

    render_period_ms = render_ms - trimesh_ms
    'If render_period_ms < 0.001 Then render_period_ms = 0.001
    'Print Using "draw time #.###"; render_period_ms
    'Print "Pixels Drawn"; Pixels_Drawn_This_Frame
    'Print "Pixels/Second"; Int(Pixels_Drawn_This_Frame / render_period_ms)

    'For tri = 0 To frame_advance
    '    Print Using "#.### "; frame_ts(tri);
    'Next tri
    'Print " "
    'For tri = 0 To frame_advance
    '    Print Using "##### "; frame_early_polls(tri);
    'Next tri


    _Limit 60
    _Display

    $Checking:On
    vCameraPsn = vCameraPsnNext

    ' keyboard polling for camera movement
    frametimestamp_now_ms = Timer(0.001)
    If frametimestamp_now_ms - frametimestamp_prior_ms < 0.0 Then
        ' timer rollover
        ' without over-analyzing just use the previous delta, even if it is somewhat wrong it is a better guess than 0.
        frametimestamp_prior_ms = frametimestamp_now_ms - frametimestamp_delta_ms
    Else
        frametimestamp_delta_ms = frametimestamp_now_ms - frametimestamp_prior_ms
    End If
    While frametimestamp_delta_ms > frametime_fullframethreshold_ms
        frametimestamp_delta_ms = frametimestamp_delta_ms - frametime_fullframe_ms
        frametimestamp_prior_ms = frametimestamp_prior_ms + frametime_fullframe_ms

        frame_advance = frame_advance + 1
        CameraPoll vCameraPsn, fYaw, fPitch
    Wend ' frametime

    KeyNow = UCase$(InKey$)
    i = 1 ' avoid deadlock
    While (KeyNow <> "") And (i < 100)

        If KeyNow = "S" Then
            Animate_Spin = Not Animate_Spin
        ElseIf KeyNow = "E" Then
            Draw_Environment_Map = Not Draw_Environment_Map
        ElseIf KeyNow = "=" Or KeyNow = "+" Then
            CameraSpeedLevel = CameraSpeedLevel + 1
            If CameraSpeedLevel > CameraSpeedLevel_max Then CameraSpeedLevel = CameraSpeedLevel_max
        ElseIf KeyNow = "-" Or KeyNow = "_" Then
            CameraSpeedLevel = CameraSpeedLevel - 1
            If CameraSpeedLevel < 0 Then CameraSpeedLevel = 0
        ElseIf KeyNow = "R" Then
            vCameraPsn.x = 0.0
            vCameraPsn.y = 0.0
            vCameraPsn.z = Camera_Start_Z
            fPitch = 0.0
            fYaw = 0.0
        ElseIf KeyNow = "G" Then
            Gouraud_Shading_Selection = Gouraud_Shading_Selection + 1
            If Gouraud_Shading_Selection > 1 Then Gouraud_Shading_Selection = 0
        ElseIf KeyNow = "B" Then
            Draw_Backface = Not Draw_Backface
        ElseIf KeyNow = "J" Then
            Jog_Motion_Selection = Jog_Motion_Selection + 1
            If Jog_Motion_Selection > 3 Then Jog_Motion_Selection = 1
        ElseIf KeyNow = "O" Then
            Jog_Motion_Selection = 0
            Animate_Spin = -1
        ElseIf KeyNow = "A" Then
            Alpha_Mask_Selection = Alpha_Mask_Selection + 1
            If Alpha_Mask_Selection > 2 Then Alpha_Mask_Selection = 0
        ElseIf Asc(KeyNow) = 27 Then
            ExitCode = 1
        End If
        KeyNow = UCase$(InKey$)
        i = i + 1
    Wend

    ' overrides
    If Vtx_Normals_Count = 0 Then Gouraud_Shading_Selection = 0

Loop Until ExitCode <> 0

For refIndex = 5 To 0 Step -1
    _FreeImage SkyBoxRef(refIndex)
Next refIndex

End
$Checking:Off

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
Data 0,0,1,0,0,1
Data 4

Data +10,+10,+10
Data +10,-10,+10
Data -10,-10,+10
Data 1,0,1,1,0,1
Data 4

' RIGHT X+
Data +10,+10,+10
Data +10,+10,-10
Data +10,-10,+10
Data 0,0,1,0,0,1
Data 0

Data +10,+10,-10
Data +10,-10,-10
Data +10,-10,+10
Data 1,0,1,1,0,1
Data 0

' LEFT X-
Data -10,+10,-10
Data -10,+10,+10
Data -10,-10,-10
Data 0,0,1,0,0,1
Data 1

Data -10,+10,+10
Data -10,-10,+10
Data -10,-10,-10
Data 1,0,1,1,0,1
Data 1

' BACK Z-
Data +10,+10,-10
Data -10,+10,-10
Data +10,-10,-10
Data 0,0,1,0,0,1
Data 5

Data -10,+10,-10
Data -10,-10,-10
Data +10,-10,-10
Data 1,0,1,1,0,1
Data 5

' TOP Y+
Data -10,+10,-10
Data +10,+10,-10
Data -10,+10,+10
Data 0,0,1,0,0,1
Data 2

Data +10,+10,-10
Data +10,+10,+10
Data -10,+10,+10
Data 1,0,1,1,0,1
Data 2

' BOTTOM Y-
Data -10,-10,+10
Data 10,-10,+10
Data -10,-10,-10
Data 0,0,1,0,0,1
Data 3

Data +10,-10,+10
Data +10,-10,-10
Data -10,-10,-10
Data 1,0,1,1,0,1
Data 3

$Checking:On
Sub CameraPoll (camloc As vec3d, yaw As Single, pitch As Single)
    Static cammove As vec3d

    If _KeyDown(32) Then
        ' Spacebar
        camloc.y = camloc.y + CameraSpeedLookup(CameraSpeedLevel) / 2
    End If

    If _KeyDown(118) Or _KeyDown(86) Then
        'V
        camloc.y = camloc.y - CameraSpeedLookup(CameraSpeedLevel) / 2
    End If

    If _KeyDown(19712) Then
        ' Right arrow
        yaw = yaw - 1.2
    End If

    If _KeyDown(19200) Then
        ' Left arrow
        yaw = yaw + 1.2
    End If

    ' forward camera movement vector
    Matrix4_MakeRotation_Y yaw, matCameraRot()
    Multiply_Vector3_Matrix4 vCameraHomeFwd, matCameraRot(), cammove
    Vector3_Mul cammove, CameraSpeedLookup(CameraSpeedLevel), cammove

    If _KeyDown(18432) Then
        ' Up arrow
        Vector3_Add camloc, cammove, camloc
    End If

    If _KeyDown(20480) Then
        ' Down arrow
        Vector3_Delta camloc, cammove, camloc
    End If

    If _KeyDown(122) Or _KeyDown(90) Then
        ' Z
        pitch = pitch + 1.0
        If pitch > 85.0 Then pitch = 85.0
    End If

    If _KeyDown(113) Or _KeyDown(81) Then
        ' Q
        pitch = pitch - 1.0
        If pitch < -85.0 Then pitch = -85.0
    End If
End Sub

Sub InsertCatalogTexture (thefile As String, h As Long)
    ' returns -1 in h if texture cannot be loaded
    Dim i As Integer
    If TextureCatalog_nextIndex > UBound(TextureCatalog) Then
        Print "need to increase texture catalog size"
        End
    End If

    For i = 0 To TextureCatalog_nextIndex
        If TextureCatalog(i).textName = thefile Then
            'match
            'Print i; " match "; thefile; TextureCatalog(i).imageHandle
            h = TextureCatalog(i).imageHandle
            '_Delay 2
            Exit Sub
        End If
    Next i

    h = _LoadImage(Obj_Directory + thefile, 32)
    If h = -1 Then
        Print "texture not loaded: "; thefile
        Exit Sub
    End If

    TextureCatalog(TextureCatalog_nextIndex).imageHandle = h
    TextureCatalog(TextureCatalog_nextIndex).textName = thefile
    TextureCatalog_nextIndex = TextureCatalog_nextIndex + 1
End Sub

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
    If totalMaterialLibrary > 0 Then Print materialFile
    'Dim temp$
    'Input "waiting..."; temp$
End Sub

Sub LoadMesh (thefile As String, tris() As mesh_triangle, indexTri As Long, v() As vec3d, leVertexTexelList As Long, vn() As vec3d, mats() As newmtl_type)
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
    'Dim VertexList(leVertexList) As vec3d ' this is now global
    Dim TexelCoord(leVertexTexelList) As vec3d ' this gets tossed after loading mesh()

    totalVertex = 0 ' refers to v() array
    totalVertexNormals = 0 ' refers to vn() array

    lineCount = 0
    lineLength = 0

    Dim rawString As String
    Dim parameter As String

    If _FileExists(thefile) = 0 Then
        Print "The object file was not found"
        Exit Sub
    End If
    Open thefile For Binary As #2
    If LOF(2) = 0 Then
        Print "The object file is empty"
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
                v(totalVertex).x = ParameterStorage(1, 0)
                v(totalVertex).y = ParameterStorage(2, 0)
                v(totalVertex).z = -ParameterStorage(3, 0) ' in obj files, Z+ points toward viewer
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
                vn(totalVertexNormals).z = -ParameterStorage(3, 0) ' in obj files, Z+ points toward viewer
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
                TexelCoord(totalVertexTexels).x = ParameterStorage(1, 0) ' it is amazing how bad at following the standard people are.
                TexelCoord(totalVertexTexels).y = ParameterStorage(2, 0) ' it is only supposed to range from 0.0 to 1.0
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

    Dim i0 As Long
    Dim i1 As Long
    Dim i2 As Long
    Dim i3 As Long
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

            While lineCursor <= lineLength
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
                i0 = FindVertexNumAbsOrRel(ParameterStorage(1, 0), mostRecentVertex)
                i1 = FindVertexNumAbsOrRel(ParameterStorage(2, 0), mostRecentVertex)
                i2 = FindVertexNumAbsOrRel(ParameterStorage(3, 0), mostRecentVertex)

                indexTri = indexTri + 1
                tris(indexTri).i0 = i0
                tris(indexTri).i1 = i1
                tris(indexTri).i2 = i2
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
                i3 = FindVertexNumAbsOrRel(ParameterStorage(4, 0), mostRecentVertex)

                indexTri = indexTri + 1
                tris(indexTri).i0 = i0
                tris(indexTri).i1 = i2
                tris(indexTri).i2 = i3
                tris(indexTri).options = tris(indexTri - 1).options
                tris(indexTri).material = useMaterialNumber

                If paramSubindex >= 1 Then
                    If (tris(indexTri).options And T1_option_no_T1) = 0 Then
                        tris(indexTri).u0 = TexelCoord(tex0).x
                        tris(indexTri).v0 = TexelCoord(tex0).y

                        tris(indexTri).u1 = TexelCoord(tex2).x
                        tris(indexTri).v1 = TexelCoord(tex2).y

                        tex3 = FindVertexNumAbsOrRel(ParameterStorage(4, 1), mostRecentVertexTexel)
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

Sub LoadMaterialFile (theFile As String, mats() As newmtl_type, totalMaterials As Long, hiddenMaterials As Long)
    Dim ParameterStorageNumeric(40) As Double
    Dim ParameterStorageText(40) As String

    Dim lineCount As Long
    Dim lineLength As Integer
    Dim lineCursor As Integer
    Dim parameterIndex As Integer
    Dim subindex As Integer
    Dim substringStart As Integer
    Dim substringLength As Integer

    Dim rawString As String
    Dim trimString As String
    Dim parameter As String
    Dim textureFilename As String

    totalMaterials = 0
    hiddenMaterials = 0
    lineCount = 0

    If _FileExists(theFile) = 0 Then
        Print "The mtl file was not found"
        Exit Sub
    End If
    Open theFile For Binary As #2
    If LOF(2) = 0 Then
        Print "The mtl file is empty"
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
            mats(totalMaterials).illum = 2
            mats(totalMaterials).Ns = 10.0
            mats(totalMaterials).options = 0

        ElseIf Left$(trimString, 6) = "map_d " Then
            ' assume that this texture has see through parts, like a fence or ironwork.
            mats(totalMaterials).options = mats(totalMaterials).options Or T1_option_no_backface_cull

        ElseIf Left$(trimString, 7) = "map_Kd " Then
            ' this is the most common texture map from an image file
            lineCursor = 8
            parameterIndex = 0

            While lineCursor < lineLength
                substringLength = 0

                ' eat spaces
                While Asc(trimString, lineCursor) <= 32
                    lineCursor = lineCursor + 1
                    If lineCursor > lineLength Then GoTo BAIL_MAP_KD
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
                    ParameterStorageText(parameterIndex) = _Trim$(parameter$)
                    'Print parameterIndex, substringStart, substringLength, "["; parameter$; "]"
                End If
            Wend

            BAIL_MAP_KD:
            If parameterIndex > 0 Then
                ' -options string comparison
                subindex = 1
                While subindex < (parameterIndex - 1)
                    If ParameterStorageText(subindex) = "-clamp" Then
                        If ParameterStorageText(subindex + 1) = "on" Then
                            mats(totalMaterials).options = mats(totalMaterials).options Or T1_option_clamp_height Or T1_option_clamp_width
                            subindex = subindex + 1
                        ElseIf ParameterStorageText(subindex + 1) = "u" Then
                            mats(totalMaterials).options = mats(totalMaterials).options Or T1_option_clamp_width
                            subindex = subindex + 1
                        ElseIf ParameterStorageText(subindex + 1) = "v" Then
                            mats(totalMaterials).options = mats(totalMaterials).options Or T1_option_clamp_height
                            subindex = subindex + 1
                        End If
                        subindex = subindex + 1
                    End If
                Wend

                ' the filename is always last in the list
                textureFilename = ParameterStorageText(parameterIndex)
                InsertCatalogTexture textureFilename, mats(totalMaterials).map_Kd
                ' check for modelers giving bad Kd values when using a texture
                If (mats(totalMaterials).Kd_r < oneOver255) And (mats(totalMaterials).Kd_g < oneOver255) And (mats(totalMaterials).Kd_b < oneOver255) Then
                    Print "Warning: Kd RGB values are zero. Texture would be all black. Changing to 1.0"
                    mats(totalMaterials).Kd_r = 1.0
                    mats(totalMaterials).Kd_g = 1.0
                    mats(totalMaterials).Kd_b = 1.0
                End If
            End If

        ElseIf Left$(trimString, 3) = "Kd " Then
            ' Diffuse color is the most dominant color
            lineCursor = 4
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
                    ParameterStorageNumeric(parameterIndex) = Val(parameter$)
                    'Print parameterIndex, substringStart, substringLength, "["; parameter$; "]"
                End If
            Wend

            BAIL_KD:
            If parameterIndex = 3 Then
                mats(totalMaterials).Kd_r = ParameterStorageNumeric(1)
                mats(totalMaterials).Kd_g = ParameterStorageNumeric(2)
                mats(totalMaterials).Kd_b = ParameterStorageNumeric(3)
            End If
            Print totalMaterials, "["; Mid$(trimString, 4); "]"

        ElseIf Left$(trimString, 3) = "Ks " Then
            ' Specular color
            lineCursor = 4
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
                    ParameterStorageNumeric(parameterIndex) = Val(parameter$)
                    'Print parameterIndex, substringStart, substringLength, "["; parameter$; "]"
                End If
            Wend

            BAIL_KS:
            If parameterIndex = 3 Then
                mats(totalMaterials).Ks_r = ParameterStorageNumeric(1)
                mats(totalMaterials).Ks_g = ParameterStorageNumeric(2)
                mats(totalMaterials).Ks_b = ParameterStorageNumeric(3)
            End If
            'Print totalMaterials, "["; Mid$(trimString, 4); "]"

        ElseIf Left$(trimString, 2) = "d " Then
            ' diaphaneity
            mats(totalMaterials).diaphaneity = Val(Mid$(trimString, 3))
            If mats(totalMaterials).diaphaneity = 0.0 Then
                hiddenMaterials = hiddenMaterials + 1
            End If

        ElseIf Left$(trimString, 3) = "Ns " Then
            ' specular power exponent
            mats(totalMaterials).Ns = Val(Mid$(trimString, 4))

        ElseIf Left$(trimString, 6) = "illum " Then
            ' illumination model
            mats(totalMaterials).illum = Int(Val(Mid$(trimString, 7)))

            ' correct overly ambitious raytracing illumination models to what we can render
            If mats(totalMaterials).illum = illum_model_reflection_map_raytrace Then mats(totalMaterials).illum = illum_model_blinn_phong
            If mats(totalMaterials).illum = illum_model_glass_map_raytrace Then mats(totalMaterials).illum = illum_model_blinn_phong
            If mats(totalMaterials).illum = illum_model_fresnel Then mats(totalMaterials).illum = illum_model_blinn_phong
            If mats(totalMaterials).illum = illum_model_refraction Then mats(totalMaterials).illum = illum_model_blinn_phong
            If mats(totalMaterials).illum = illum_model_fresnel_refraction Then mats(totalMaterials).illum = illum_model_blinn_phong
            If mats(totalMaterials).illum = illum_model_glass_map Then mats(totalMaterials).illum = illum_model_reflection_map

        End If
        BAIL_INDENT_MATERIAL:
    Loop
    Close #2

    Print totalMaterials, "Total Materials"
    Print "#", "Name", "Kd_r", "Kd_g", "Kd_b", "d", "Ns", "illum"
    Dim i As Long
    For i = 1 To totalMaterials
        Print i, mats(i).textName, mats(i).Kd_r, mats(i).Kd_g, mats(i).Kd_b, mats(i).diaphaneity, mats(i).Ns, mats(i).illum
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
    'mag = -2.0 * Vector3_DotProduct!(i, normal)
    'Vector3_Mul normal, mag, bounce
    'Vector3_Add bounce, i, o
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
            ' v (0 to 1) goes from -y to +y
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
        ' v (0 to 1) goes from -y to +y
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
    m(3, 0) = 0.0: m(3, 1) = 0.0: m(3, 2) = 0.0: m(3, 3) = 1.0
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
    m(2, 0) = 0.0: m(2, 1) = 0.0: m(2, 2) = 1.0: m(2, 3) = 0.0
    m(3, 0) = 0.0: m(3, 1) = 0.0: m(3, 2) = 0.0: m(3, 3) = 1.0
End Sub

Sub Matrix4_MakeRotation_Y (deg As Single, m( 3 , 3) As Single)
    m(0, 0) = Cos(_D2R(deg))
    m(0, 1) = 0.0
    m(0, 2) = Sin(_D2R(deg))
    m(0, 3) = 0.0
    m(1, 0) = 0.0: m(1, 1) = 1.0: m(1, 2) = 0.0: m(1, 3) = 0.0
    m(2, 0) = -Sin(_D2R(deg))
    m(2, 1) = 0.0
    m(2, 2) = Cos(_D2R(deg))
    m(2, 3) = 0.0
    m(3, 0) = 0.0: m(3, 1) = 0.0: m(3, 2) = 0.0: m(3, 3) = 1.0
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
    m(3, 0) = 0.0: m(3, 1) = 0.0: m(3, 2) = 0.0: m(3, 3) = 1.0
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

' This is the generic 3D point projection formula with a 4 by 4 matrix.
'  input w is assumed always 1, so you can send a vec3d.
' Practical reasons for this existing are to:
'  Simulate lens field of view angles.
'  Normalize depth to maximum numeric range between the near and far plane.
'  Have the ability to swap coordinates. For example swap input Y and Z.
'  Give a 4th channel with which to perform perspective division (or not).
'  Perform non-perspective projections like isometric.
'  Warp projections to depict momentary disorientation from a large blast.
Sub ProjectMatrixVector4 (i As vec3d, m( 3 , 3) As Single, o As vec4d)
    Dim www As Single
    o.x = i.x * m(0, 0) + i.y * m(1, 0) + i.z * m(2, 0) + m(3, 0)
    o.y = i.x * m(0, 1) + i.y * m(1, 1) + i.z * m(2, 1) + m(3, 1)
    o.z = i.x * m(0, 2) + i.y * m(1, 2) + i.z * m(2, 2) + m(3, 2)
    www = i.x * m(0, 3) + i.y * m(1, 3) + i.z * m(2, 3) + m(3, 3)

    ' Normalizing
    If www <> 0.0 Then
        o.w = 1 / www ' optimization
        o.x = o.x * o.w
        o.y = o.y * o.w
    End If
End Sub

' Simplified Perspective Projection Only
' The concept of perspective projection being division by z gets lost in the matrix.
' Here is the equation from above, rewritten without the always zero and unused Z matrix element values.
' o.z is unused.
Sub ProjectPerspectiveVector4 (i As vec3d, m( 3 , 3) As Single, o As vec4d)
    If i.z <> 0.0 Then
        o.w = 1.0 / i.z ' we need 1/z for later, plus we can multiply it with x and y instead of divide.
        o.x = i.x * m(0, 0) * o.w ' same as i.x * m(0,0) / i.z
        o.y = i.y * m(1, 1) * o.w
    Else
        o.w = 0.0
    End If
End Sub


Sub VertexColorAlphaTriangle (A As vertex9, B As vertex9, C As vertex9)
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

    ' caching of 4 texels in bilinear mode
    Static T1_last_cache As _Unsigned Long
    T1_last_cache = &HFFFFFFFF ' Invalidate texel cache

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

            ' metrics
            If col < draw_max_x Then Pixels_Drawn_This_Frame = Pixels_Drawn_This_Frame + (draw_max_x - col)

            ' Draw the Horizontal Scanline
            ' Optimization: before entering this loop, must have done tex_z = 1 / tex_w
            work_address = work_row_base + 4 * col
            zbuf_index = row * Size_Render_X + col
            While col < draw_max_x

                If Screen_Z_Buffer(zbuf_index) = 0.0 Or tex_z < Screen_Z_Buffer(zbuf_index) Then
                    If (Texture_options And T1_option_no_Z_write) = 0 Then
                        Screen_Z_Buffer(zbuf_index) = tex_z + Z_Fight_Bias
                    End If

                    Static pixel_combine As _Unsigned Long
                    pixel_combine = _RGB32(tex_r, tex_g, tex_b)

                    Static pixel_existing As _Unsigned Long
                    pixel_alpha = T1_mod_A
                    If pixel_alpha < 0.998 Then
                        pixel_existing = _MemGet(work_mem_info, work_address, _Unsigned Long)
                        pixel_combine = _RGB32((  _red32(pixel_combine) - _Red32(pixel_existing))   * pixel_alpha + _red32(pixel_existing), _
                                               (_green32(pixel_combine) - _Green32(pixel_existing)) * pixel_alpha + _green32(pixel_existing), _
                                               ( _Blue32(pixel_combine) - _Blue32(pixel_existing))  * pixel_alpha + _blue32(pixel_existing))

                        ' x = (p1 - p0) * ratio + p0 is equivalent to
                        ' x = (1.0 - ratio) * p0 + ratio * p1
                    End If
                    _MemPut work_mem_info, work_address, pixel_combine

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

Sub TextureWithAlphaTriangle (A As vertex9, B As vertex9, C As vertex9)
    ' is able to handle non power of 2 texture sizes
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
    Static pixel_value As _Unsigned Long ' The ARGB value to write to screen

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

    ' caching of 4 texels in bilinear mode
    Static T1_last_cache As _Unsigned Long
    T1_last_cache = &HFFFFFFFF ' invalidate

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

            ' metrics
            If col < draw_max_x Then Pixels_Drawn_This_Frame = Pixels_Drawn_This_Frame + (draw_max_x - col)

            ' Draw the Horizontal Scanline
            ' Optimization: before entering this loop, must have done tex_z = 1 / tex_w
            ' Relies on some shared T1 variables over by Texture1
            work_address = work_row_base + 4 * col
            zbuf_index = row * Size_Render_X + col
            While col < draw_max_x

                ' Check Z-Buffer early to see if we even need texture lookup and color combine
                ' Note: Only solid (non-transparent) pixels update the Z-buffer
                If Screen_Z_Buffer(zbuf_index) = 0.0 Or tex_z < Screen_Z_Buffer(zbuf_index) Then

                    ' Read Texel
                    ' Relies on shared T1_ variables
                    Static cc As Long
                    Static ccp As Long
                    Static rr As Long
                    Static rrp As Long

                    Static cm5 As Single
                    Static rm5 As Single

                    ' Recover U and V
                    ' Offset so the transition appears in the center of an enlarged texel instead of a corner.
                    cm5 = (tex_u * tex_z) - 0.5
                    rm5 = (tex_v * tex_z) - 0.5

                    If Texture_options And T1_option_clamp_width Then
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
                    Else
                        ' tile
                        ' positive modulus
                        cc = Int(cm5) Mod T1_width
                        If cc < 0 Then cc = cc + T1_width

                        ccp = cc + 1
                        If ccp > T1_width_MASK Then ccp = 0
                    End If

                    If Texture_options And T1_option_clamp_height Then
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
                    Else
                        ' tile
                        ' positive modulus
                        rr = Int(rm5) Mod T1_height
                        If rr < 0 Then rr = rr + T1_height

                        rrp = (rr + 1)
                        If rrp > T1_height_MASK Then rrp = 0
                    End If

                    ' 4 point bilinear temp vars
                    Static Frac_cc1_FIX7 As Integer
                    Static Frac_rr1_FIX7 As Integer
                    ' 0 1
                    ' . .
                    Static bi_r0 As Integer
                    Static bi_g0 As Integer
                    Static bi_b0 As Integer
                    Static bi_a0 As Integer
                    ' . .
                    ' 2 3
                    Static bi_r1 As Integer
                    Static bi_g1 As Integer
                    Static bi_b1 As Integer
                    Static bi_a1 As Integer

                    ' color blending
                    Static a0 As Integer

                    Frac_cc1_FIX7 = (cm5 - Int(cm5)) * 128
                    Frac_rr1_FIX7 = (rm5 - Int(rm5)) * 128

                    ' Caching of 4 texels
                    Static T1_this_cache As _Unsigned Long
                    Static T1_uv_0_0 As Long
                    Static T1_uv_1_0 As Long
                    Static T1_uv_0_1 As Long
                    Static T1_uv_1_1 As Long

                    T1_this_cache = _ShL(rr, 16) Or cc
                    If T1_this_cache <> T1_last_cache Then

                        _MemGet T1_mblock, T1_mblock.OFFSET + (cc + rr * T1_width) * 4, T1_uv_0_0
                        _MemGet T1_mblock, T1_mblock.OFFSET + (ccp + rr * T1_width) * 4, T1_uv_1_0
                        _MemGet T1_mblock, T1_mblock.OFFSET + (cc + rrp * T1_width) * 4, T1_uv_0_1
                        _MemGet T1_mblock, T1_mblock.OFFSET + (ccp + rrp * T1_width) * 4, T1_uv_1_1

                        T1_last_cache = T1_this_cache
                    End If

                    If Texture_options And T1_option_alpha_hard_edge Then
                        ' hard edge
                        ' due to offset by -0.5, the bottom-right texel is the mask.
                        ' ___ ___
                        '|  _ _  |
                        '| |_|_| |
                        '| |_|m| |
                        '|___ ___|
                        '
                        a0 = _Alpha32(T1_uv_1_1)
                        If a0 < T1_Alpha_Threshold Then a0 = 0
                    Else
                        ' soft edge
                        ' mask is bilinear blend Alpha channel from texture
                        bi_a0 = _Alpha32(T1_uv_0_0)
                        bi_a0 = _ShR((_Alpha32(T1_uv_1_0) - bi_a0) * Frac_cc1_FIX7, 7) + bi_a0

                        bi_a1 = _Alpha32(T1_uv_0_1)
                        bi_a1 = _ShR((_Alpha32(T1_uv_1_1) - bi_a1) * Frac_cc1_FIX7, 7) + bi_a1

                        a0 = _ShR((bi_a1 - bi_a0) * Frac_rr1_FIX7, 7) + bi_a0
                    End If


                    If a0 > 0 Then
                        ' Color Combiner math for Alpha
                        If Texture_options And T1_option_alpha_single_bit Then
                            a0 = T1_mod_A * 255.0
                        Else
                            a0 = a0 * T1_mod_A
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

                        ' Color Combiner math for RGB
                        pixel_value = _RGB32(_Red32(pixel_value) * tex_r, _Green32(pixel_value) * tex_g, _Blue32(pixel_value) * tex_b)

                        If a0 < 255 Then
                            ' Alpha blend
                            Static pixel_existing As _Unsigned Long
                            Static pixel_alpha As Single
                            pixel_alpha = a0 * oneOver255
                            pixel_existing = _MemGet(work_mem_info, work_address, _Unsigned Long)

                            pixel_value = _RGB32((  _red32(pixel_value) -  _Red32(pixel_existing))  * pixel_alpha +   _red32(pixel_existing), _
                                                 (_green32(pixel_value) - _Green32(pixel_existing)) * pixel_alpha + _green32(pixel_existing), _
                                                 ( _Blue32(pixel_value) - _Blue32(pixel_existing))  * pixel_alpha +  _blue32(pixel_existing))

                            _MemPut work_mem_info, work_address, pixel_value

                            If (Texture_options And T1_option_no_Z_write) = 0 Then
                                If a0 >= T1_Alpha_Threshold Then Screen_Z_Buffer(zbuf_index) = tex_z + Z_Fight_Bias
                            End If
                        Else
                            ' Solid
                            _MemPut work_mem_info, work_address, pixel_value
                            Screen_Z_Buffer(zbuf_index) = tex_z + Z_Fight_Bias
                        End If

                    End If ' a0
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
    ' Texture_options is ignored, Z depth is ignored.
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

            ' metrics
            If col < draw_max_x Then Pixels_Drawn_This_Frame = Pixels_Drawn_This_Frame + (draw_max_x - col)

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

                T1_this_cache = _ShL(rr, 16) Or cc
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

    ' caching of 4 texels in bilinear mode
    Static T1_last_cache As _Unsigned Long
    T1_last_cache = &HFFFFFFFF ' invalidate

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

            ' metrics
            If col < draw_max_x Then Pixels_Drawn_This_Frame = Pixels_Drawn_This_Frame + (draw_max_x - col)

            ' Draw the Horizontal Scanline
            ' Optimization: before entering this loop, must have done tex_z = 1 / tex_w
            work_address = work_row_base + 4 * col
            zbuf_index = row * Size_Render_X + col
            While col < draw_max_x

                If Screen_Z_Buffer(zbuf_index) = 0.0 Or tex_z < Screen_Z_Buffer(zbuf_index) Then
                    If (Texture_options And T1_option_no_Z_write) = 0 Then
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
                    pixel_combine = _RGB32(r0 * T1_mod_R, g0 * T1_mod_G, b0 * T1_mod_B)

                    Static pixel_existing As _Unsigned Long
                    pixel_alpha = T1_mod_A
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
