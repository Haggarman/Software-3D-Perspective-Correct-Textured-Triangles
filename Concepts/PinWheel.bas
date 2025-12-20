Option _Explicit
_Title "PinWheel 4: Move with arrows, scale with + or - , F to change filter, S to spin"
' 2025 Haggarman
'
' The point is to look for an all green pinwheel. Gaps or overdraw will show up as other colors.
' Press F to flip to the actual rendering and again to go back to showing overdraw.
'
Dim Shared DISP_IMAGE As Long
Dim Shared WORK_IMAGE As Long
Dim Shared FILTER_IMAGE As Long
Dim Shared Size_Screen_X As Integer, Size_Screen_Y As Integer
Dim Shared Size_Render_X As Integer, Size_Render_Y As Integer

' MODIFY THESE if you want.
Size_Screen_X = 800
Size_Screen_Y = 600
Size_Render_X = Size_Screen_X \ 6 ' render size
Size_Render_Y = Size_Screen_Y \ 6 ' think of it as the number of screen pixels that one rendered pixel takes

DISP_IMAGE = _NewImage(Size_Screen_X, Size_Screen_Y, 32)
Screen DISP_IMAGE
_DisplayOrder _Software

WORK_IMAGE = _NewImage(Size_Render_X, Size_Render_Y, 32)
FILTER_IMAGE = _NewImage(Size_Render_X, Size_Render_Y, 32)
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
End Type

Type vertex5
    x As Single
    y As Single
    w As Single
    u As Single
    v As Single
End Type

' Viewing area clipping
Dim Shared clip_min_y As Long, clip_max_y As Long
Dim Shared clip_min_x As Long, clip_max_x As Long
clip_min_y = 0
clip_max_y = Size_Render_Y - 1
clip_min_x = 0
clip_max_x = Size_Render_X ' not (-1) because rounding rule drops one pixel on right

' Fog color is used as the background color
Dim Shared Fog_color As Long
Fog_color = _RGB32(47, 78, 105)

' Load textures from file
Dim Shared TextureCatalog_lastIndex As Integer
TextureCatalog_lastIndex = 0
Dim Shared TextureCatalog(TextureCatalog_lastIndex) As Long
TextureCatalog(0) = _LoadImage("Origin16x16.png", 32)

' Error _LoadImage returns -1 as an invalid handle if it cannot load the image.
Dim refIndex As Integer
For refIndex = 0 To TextureCatalog_lastIndex
    If TextureCatalog(refIndex) = -1 Then
        Print "Could not load texture file for index: "; refIndex
        End
    End If
Next refIndex

' These T1 Texture characteristics are read later on during drawing.
Dim Shared T1_CatalogIndex As Long
Dim Shared T1_ImageHandle As Long
Dim Shared T1_width As Integer, T1_width_MASK As Integer
Dim Shared T1_height As Integer, T1_height_MASK As Integer
Dim Shared T1_mblock As _MEM
' Optimization requires that width and height be powers of 2.
'  That means: 2,4,8,16,32,64,128,256...
T1_SelectCatalogTexture 0


' Load the object mesh
' Load what is called a mesh from data statements.
' (x0,y0,z0) (x1,y1,z1) (x2,y2,z2) (u0,v0) (u1,v1) (u2,v2) texture options
Dim Triangles_In_The_Object As Integer
Triangles_In_The_Object = 8

Dim Shared Mesh_Last_Element As Integer
Mesh_Last_Element = Triangles_In_The_Object - 1
Dim mesh(Mesh_Last_Element) As triangle

Dim tri_num As Integer
Restore MESHDATA

For tri_num = 0 To Mesh_Last_Element
    Read mesh(tri_num).x0
    Read mesh(tri_num).y0
    Read mesh(tri_num).z0

    Read mesh(tri_num).x1
    Read mesh(tri_num).y1
    Read mesh(tri_num).z1

    Read mesh(tri_num).x2
    Read mesh(tri_num).y2
    Read mesh(tri_num).z2

    Read mesh(tri_num).u0
    Read mesh(tri_num).v0
    Read mesh(tri_num).u1
    Read mesh(tri_num).v1
    Read mesh(tri_num).u2
    Read mesh(tri_num).v2

    Read mesh(tri_num).texture
    Read mesh(tri_num).options
Next tri_num

' Screen Scaling
Dim halfWidth As Single
Dim halfHeight As Single
halfWidth = Size_Render_X / 2
halfHeight = Size_Render_Y / 2

' Triangle Vertex List
Dim vertexA As vertex5
Dim vertexB As vertex5
Dim vertexC As vertex5

' code execution time
Dim start_ms As Double
Dim finish_ms As Double

' Rotation
Dim matRotZ(3, 3) As Single
Dim point0 As vec3d
Dim point1 As vec3d
Dim point2 As vec3d

Dim pointRotZ0 As vec3d
Dim pointRotZ1 As vec3d
Dim pointRotZ2 As vec3d

' This is so that the object animates by rotating
Dim spinAngleDegZ As Single
spinAngleDegZ = 0.0

' Main loop stuff
Dim KeyNow As String
Dim ExitCode As Integer
Dim Animate_Spin As Integer
Dim Shared Filter_Selection As Integer
Dim FrameCounter As Long
Dim animationStep As Integer
Dim depth_w As Single
Dim myscale As Single
Dim pin As vec3d

Dim frame_advance As Integer

main:
ExitCode = 0
Animate_Spin = -1
Filter_Selection = 1
FrameCounter = 0
animationStep = 0
myscale = 32 ' scale up a "unit triangle" that is otherwise 0..1

' screen pin location of the triangles
pin.x = halfWidth
pin.y = halfHeight
pin.z = 1.0 ' cannot be 0

frame_advance = 1

Do
    start_ms = Timer(.001)
    _Dest WORK_IMAGE
    Cls , Fog_color
    _Source WORK_IMAGE

    ' Erase Screen_Z_Buffer
    ' This is a qbasic only optimization. it sets the value to zero. it saves 10 ms.
    ReDim Screen_Z_Buffer(Screen_Z_Buffer_MaxElement)

    If Animate_Spin Then
        spinAngleDegZ = spinAngleDegZ + frame_advance * 0.460
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

    T1_SelectCatalogTexture 0 ' choose the first loaded texture and set the T1_ texture vars

    ' Normally we would project the triangle points from 3D to 2D.
    ' Here we fill in directly from the mesh, scaling X and Y by myscale.
    depth_w = 1.0 / pin.z
    For tri_num = 0 To Mesh_Last_Element
        If tri_num = animationStep Then _Continue

        point0.x = mesh(tri_num).x0
        point0.y = mesh(tri_num).y0
        point0.z = mesh(tri_num).z0

        point1.x = mesh(tri_num).x1
        point1.y = mesh(tri_num).y1
        point1.z = mesh(tri_num).z1

        point2.x = mesh(tri_num).x2
        point2.y = mesh(tri_num).y2
        point2.z = mesh(tri_num).z2

        ' Rotate in Z-Axis
        Multiply_Vector3_Matrix4 point0, matRotZ(), pointRotZ0
        Multiply_Vector3_Matrix4 point1, matRotZ(), pointRotZ1
        Multiply_Vector3_Matrix4 point2, matRotZ(), pointRotZ2

        vertexA.x = pin.x + pointRotZ0.x * myscale
        vertexA.y = pin.y + pointRotZ0.y * myscale
        vertexA.w = depth_w
        vertexA.u = mesh(tri_num).u0 * T1_width * depth_w
        vertexA.v = mesh(tri_num).v0 * T1_height * depth_w

        vertexB.x = pin.x + pointRotZ1.x * myscale
        vertexB.y = pin.y + pointRotZ1.y * myscale
        vertexB.w = depth_w
        vertexB.u = mesh(tri_num).u1 * T1_width * depth_w
        vertexB.v = mesh(tri_num).v1 * T1_height * depth_w

        vertexC.x = pin.x + pointRotZ2.x * myscale
        vertexC.y = pin.y + pointRotZ2.y * myscale
        vertexC.w = depth_w
        vertexC.u = mesh(tri_num).u2 * T1_width * depth_w
        vertexC.v = mesh(tri_num).v2 * T1_height * depth_w

        TexturedNonlitTriangle vertexA, vertexB, vertexC
    Next tri_num


    If Filter_Selection = 0 Then
        _PutImage , WORK_IMAGE, DISP_IMAGE
        _Dest DISP_IMAGE
        Locate 1, 1
        Color _RGB32(177, 227, 255), Fog_color
        Print "Actual Framebuffer";
    Else
        CalculateOverdrawFilter
        _PutImage , FILTER_IMAGE, DISP_IMAGE
        _Dest DISP_IMAGE
        ' Show a Legend
        Locate 1, 1
        Color _RGB32(177, 227, 255), Fog_color
        Print "Overdraw:";
        Color 0, Fog_color
        Print " 0 ";
        Color 0, _RGB32(50, 250, 50)
        Print " 1 ";
        Color 0, _RGB32(155, 67, 222)
        Print " 2 ";
        Color 0, _RGB32(200, 194, 28)
        Print " 3+"
    End If
    finish_ms = Timer(.001)

    '_Dest DISP_IMAGE
    'Locate Int(Size_Screen_Y \ 16) - 2, 1
    'Color _RGB32(177, 227, 255), Fog_color
    'Print Using "render time #.###"; finish_ms - start_ms

    Color _RGB32(177, 227, 255), Fog_color
    If Animate_Spin Then
        Print "Press S to Stop Spin"
    Else
        Print "Press S to Start Spin"
    End If

    _Limit 60
    _Display

    FrameCounter = FrameCounter + 1
    If FrameCounter >= 50 Then
        FrameCounter = 0
        animationStep = animationStep + 1
        If animationStep > Mesh_Last_Element + 4 Then animationStep = 0
    End If

    KeyNow = UCase$(InKey$)
    If KeyNow <> "" Then

        If Asc(KeyNow) = 27 Then
            ExitCode = 1
        ElseIf KeyNow = "+" Then
            myscale = myscale + 1
        ElseIf KeyNow = "=" Then
            myscale = myscale + 1
        ElseIf KeyNow = "-" Then
            myscale = myscale - 1
        ElseIf KeyNow = "F" Then
            Filter_Selection = Filter_Selection + 1
            If Filter_Selection >= 2 Then Filter_Selection = 0
        ElseIf KeyNow = "S" Then
            Animate_Spin = Not Animate_Spin
        End If
    End If

    If _KeyDown(19712) Then
        ' Right arrow
        pin.x = pin.x + 0.721
    End If

    If _KeyDown(19200) Then
        ' Left arrow
        pin.x = pin.x - 0.721
    End If

    If _KeyDown(18432) Then
        ' Up arrow
        pin.y = pin.y - 0.721
    End If

    If _KeyDown(20480) Then
        ' Down arrow
        pin.y = pin.y + 0.721
    End If

Loop Until ExitCode <> 0


For refIndex = TextureCatalog_lastIndex To 0 Step -1
    _FreeImage TextureCatalog(refIndex)
Next refIndex

End

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
' texture_index, texture_options
' .0 = T1_option_clamp_width
' .1 = T1_option_clamp_height

' An 8 triangle pinwheel
MESHDATA:
Data 0,0,1,0,-1,1,1,-1,1
Data 0.5,0.5,0.5,0,1,0
Data 0,0

Data 0,0,1,1,-1,1,1,0,1
Data 0.5,0.5,1,0,1,0.5
Data 0,0

Data 0,0,1,1,0,1,1,1,1
Data 0.5,0.5,1,0.5,1,1
Data 0,0

Data 0,0,1,1,1,1,0,1,1
Data 0.5,0.5,1,1,0.5,1
Data 0,0

Data 0,0,1,0,1,1,-1,1,1
Data 0.5,0.5,0.5,1,0,1
Data 0,0

Data 0,0,1,-1,1,1,-1,0,1
Data 0.5,0.5,0,1,0,0.5
Data 0,0

Data 0,0,1,-1,0,1,-1,-1,1
Data 0.5,0.5,0,0.5,0,0
Data 0,0

Data 0,0,1,-1,-1,1,0,-1,1
Data 0.5,0.5,0,0,0.5,0
Data 0,0

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

Sub T1_SelectCatalogTexture (texnum As Integer)
    ' Fill in Texture 1 data
    If (texnum >= 0) And (texnum <= TextureCatalog_lastIndex) Then
        T1_CatalogIndex = texnum
    Else
        ' default
        T1_CatalogIndex = 0
    End If
    T1_ImageHandle = TextureCatalog(T1_CatalogIndex)
    T1_mblock = _MemImage(T1_ImageHandle)
    T1_width = _Width(T1_ImageHandle): T1_width_MASK = T1_width - 1
    T1_height = _Height(T1_ImageHandle): T1_height_MASK = T1_height - 1
End Sub

Sub CalculateOverdrawFilter
    ' Colorizes based on how many times pixels were overdrawn

    ' Filter Screen Memory Pointers
    Static filter_mem_info As _MEM
    Static filter_next_row_step As _Offset
    Static filter_row_base As _Offset ' Calculated every next row
    Static filter_address As _Offset ' Calculated at every starting column

    filter_mem_info = _MemImage(FILTER_IMAGE)
    filter_next_row_step = 4 * Size_Render_X

    Static R As Long
    Static C As Long
    Static zbuf_index As _Unsigned Long ' Z-Buffer
    Static lookup_color As Long
    Static pixel_value As _Unsigned Long

    _Dest FILTER_IMAGE
    filter_row_base = filter_mem_info.OFFSET
    For R = 0 To Size_Render_Y - 1
        filter_address = filter_row_base
        zbuf_index = R * Size_Render_X
        For C = 0 To Size_Render_X - 1
            lookup_color = Int(Screen_Z_Buffer(zbuf_index))
            Select Case lookup_color
                Case 0:
                    pixel_value = Fog_color
                Case 1:
                    pixel_value = _RGB32(50, 250, 50)
                Case 2:
                    pixel_value = _RGB32(155, 67, 222)
                Case Else:
                    pixel_value = _RGB32(200, 194, 28)
            End Select

            _MemPut filter_mem_info, filter_address, pixel_value
            filter_address = filter_address + 4
            zbuf_index = zbuf_index + 1
        Next C
        filter_row_base = filter_row_base + filter_next_row_step
    Next R

End Sub

Sub TexturedNonlitTriangle (A As vertex5, B As vertex5, C As vertex5)
    ' this draws texture T1 using point (nearest) sampling
    ' Required Global Vars:
    '  WORK_IMAGE, clip_min_y, clip_max_y, clip_min_x, clip_max_x, Size_Render_X,
    '  T1_mblock, T1_width, T1_width_MASK, T1_height, T1_height_MASK

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
    Static zbuf_index As _Unsigned Long ' Z-Buffer
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
            zbuf_index = row * Size_Render_X + col
            screen_address = screen_row_base + 4 * col
            While col < draw_max_x
                ' special use of Z buffer to count the number of times a pixel was written.
                Screen_Z_Buffer(zbuf_index) = Screen_Z_Buffer(zbuf_index) + 1

                Static cc As Integer
                Static rr As Integer

                ' Relies on some shared T1 variables over by Texture1
                Static T1_address_pointer As _Offset

                Static cm5 As Single
                Static rm5 As Single

                ' Recover U and V
                cm5 = (tex_u * tex_z)
                rm5 = (tex_v * tex_z)

                ' clamp
                If cm5 < 0.0 Then cm5 = 0.0
                If cm5 >= T1_width_MASK Then
                    ' 15.0 and up
                    cc = T1_width_MASK
                Else
                    ' 0 1 2 .. 13 14.999
                    cc = Int(cm5)
                End If

                ' clamp
                If rm5 < 0.0 Then rm5 = 0.0
                If rm5 >= T1_height_MASK Then
                    ' 15.0 and up
                    rr = T1_height_MASK
                Else
                    rr = Int(rm5)
                End If

                'uv_0_0 = Texture1(cc, rr)
                T1_address_pointer = T1_mblock.OFFSET + (cc + rr * T1_width) * 4
                _MemGet T1_mblock, T1_address_pointer, pixel_value

                _MemPut screen_mem_info, screen_address, pixel_value
                'PSet (col, row), pixel_value

                zbuf_index = zbuf_index + 1
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

