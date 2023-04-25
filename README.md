## Language
 This collection of BASIC programs is written for the QB64 compiler.
 https://github.com/QB64-Phoenix-Edition/
## Description
 How were 3D triangles drawn by the first PC graphics accelerators in 1997? This was my deep dive into understanding the software algorithms involved in drawing triangles line by line. This is a software simulation only approach.
 
 For extra fun there are blending operations like Fog, Vertex Color, and Directional/Ambient lighting.
## Screenshots
 ![Textured Cube Plains 640](https://user-images.githubusercontent.com/96515734/224503609-ac961d99-c086-400e-b1f9-a1a8c639f918.PNG)
 ![Vertex Color Cube](https://user-images.githubusercontent.com/96515734/219141230-6d9a9d36-ec02-4c5a-94eb-a01773a02b5b.PNG)
 ![Dither Color Cube](https://user-images.githubusercontent.com/96515734/219270641-b560848f-5ef0-429c-9c9c-6191e8ac88a9.png)
 ![Vertex Alpha](https://user-images.githubusercontent.com/96515734/225165954-80881ea9-3c7f-4ebb-9fe2-b48b3d54a161.PNG)
 ![Texture Z Fight Donut](https://user-images.githubusercontent.com/96515734/224503149-33eb6e8a-73cc-473b-a61c-5c3262b0c93f.PNG)
 ![Skybox Long Tube 800](https://user-images.githubusercontent.com/96515734/224502776-09c13ed1-60a6-4074-a7d4-dd2b05e30085.png)
 ![Skybox Treess](https://user-images.githubusercontent.com/96515734/234192793-af9a3844-cfc6-4cdb-a988-e38d5a0ae531.PNG)

## Approach
 QBASIC was used because of the very quick edit-compile-test iterations. I wanted this exploration to be fun.
 
 A verbose style of naming variables and keeping ideas separated per source code line is used.
 
 Floating point numbers are used for the sake of understanding. On early 3D accelerators there were many different fixed-point number combinations for the sake of speed and reduced transistor count. But all that bit shifting hinders understanding.
 
 I believe it would be much easier to translate this BASIC code to C-lang or Python, than it was for me to catch onto the nuances from other's example C++ code that was using STL std::list, templates and pointer tricks.
 
 No dropping to assembly or using pokes!

## Capabilities
 Let's list what has currently been implemented or explored, Final Reality Advanced benchmark style.
 
Y/N | 3D graphics options
-- | --------
Yes | Texture bi-linear filtering
Yes | Z-buffer sorting
No | Texture mip-mapping
No | Texture tri-linear mapping
Yes | Depth Fog
Yes | Specular gouraud
Yes | Vertex Alpha
No | Alpha blending (crossfade)
Yes | Additive alpha (lighten)
Yes | Multiplicative alpha (darken)
Yes | Subpixel accuracy

## Triangles
### Vertex
 The triangles are specified by vertexes A, B, and C. They are sorted by the triangle drawing subroutine so that A is always on top and C is always on the bottom. That still leaves two categories where the knee at B faces left or right. The triangle drawing subroutine also adjusts for this so that pixels are drawn from left to right.
 ![TrianglesABC](https://user-images.githubusercontent.com/96515734/220204499-62aaed3c-f1fe-4c07-9c64-1c61564219e7.PNG)
### DDA
 The DDA (Digital Difference Analyzer) algorithm is used to simultaneously step on whole number Y increments from point A to point C on the major edge, and from point A to point B on the minor edge. When the Minor Edge DDAs reach vertexBy, the start values and steps are recalculated to be from point B to point C. Note that this case also handles a flat-topped triangle where vertexBy = vertexAy.

### Pre-stepping
 The start value of Y at point A is pre-stepped ahead to the next highest integer pixel row using the ceiling (round up) function. To ensure that the sampling is visually correct, the X major, X minor, and vertex attributes (U, V, R, G, B, etc.) are also pre-stepped forward by the same amount. This prestep of Y also factors in the clipping window so that the DDA accumulators are correctly advanced to the top row of the clipping region.

### Why use DDA?
 DDA was used because not all math operations complete in the same amount of time. Dividing once before a loop and then accumulating that quotient within the loop, is going to be faster than dividing every single step within the loop. Sneakily, many of the earliest PC graphics accelerators left this initial division calculation up the main system CPU.

## Clipping
### Near Frustum Clipping
 For this discussion, I am asserting that an object moving forward from the viewer increases in +Z distance. Note this can differ in well-known graphics libraries.

 Projecting and rendering what is behind the camera (-Z) makes no sense perceptually. Although it might be mathematically correct for surfaces to invert that pass the Z=0 camera plane, we do not have double-sided eyeballs that are able to simultaneously project light onto both sides of our retinas. So we need to handle this limited field of view while rendering.

We have 3 options:

1. Constrain movement so that any Z coordinate can never be less than the near frustrum plane Z value.
2. Do not draw (as in cull) the triangle if any Z is less than the near frustum plane Z value.
3. Clip the triangle so that its Z coordinates remain at least at the near frustum.

I would say the options follow a natural progression, where many programmers give up before reaching option 3.

With option 3, we are fortunate with front clipping because the near frustum plane is always parallel to the rendering screen surface (with traditional single-point projection). This reduces the required math down from a 3D vector intersecting a 3D plane, to a 2D line intersecting a 1D plane.

#### Tesselation

We are also in a sense unlucky in trying to just draw triangles. If one Z value (of 3) needs to be clipped, this creates a 4 point "quadrangle". We need to tesselate to create 2 triangles out of this quadrangle.

#### Winding Order

Imagine hammering 3 nails partially into a board. The nails representing the vertexes of the triangle. Then proceed to wrap a string around the outline of these nails. You have 2 ways of winding: clockwise or counter-clockwise. Preserving winding order preserves which side of the tesselated triangle is facing the viewer.

Setting some ground rules can make this clipping process less painful. In this codebase, two triangles share a side (vertex A and vertex C). Triangles wind in the following order:

Triangle 1: A to B, B to C, **C to A**

Triangle 2: **A to C**, C to D, D to A

![Triangle Near Clip](https://user-images.githubusercontent.com/96515734/224221060-d9673899-82c3-4f87-b502-88fde7c958ec.png)

#### Number of Triangles

The near clipping function returns the number of triangles (n = 0, 1, or 2) after clipping.

- n = 0: The input triangle is culled (not drawn).

- n = 1: Do not need Triangle 2, so do not update vertex D but do update vertexes A, B, and C.

- n = 2: Cautiously update vertex D along with vertexes A, B, and C. Preserve "winding order".

#### Demonstration

The test program that was used to develop and debug the NearClip function is available in the concepts folder. Program *NearFrustumClipTriangleAttributes.bas* rotates and clips a single triangle, while animating the winding order as crawling dots.

### Backface culling

Imagine ink bleeding through paper so that both sides have ink on them. The printed side is the front face, and the opposite bled-through side is the back face. Seeing both sides is sometimes desirable, like for a leaf. But with closed solid objects made of multiple triangles, it is more efficient to not draw the back-facing triangles because they will never be seen.

The winding order of the triangle's vertexes determines which side is the front face. The sign (positive or negative) of the triangle's **surface normal** as compared to a normalized ray extending out from the viewer (using the dot product) can determine which side of the triangle is facing the viewer. 

If the triangle were to be viewed perfectly edge-on to have a dot product value of 0, it is also invisible because it is infinitely thin. So flagging and then not drawing the triangle if this value is less than or equal to 0.0 accomplishes backface culling.

## Texture Filters
### Texture Magnification
 The following texel filters are selectable in the examples that showcase them:
ID | Name | Description
-- | ---- | ----
0 | Nearest | Blocky sampling of a single texture point.
1 | 3-Point N64 | A distinct look using Barycentric (area) math concepts.
2 | Bilinear Fix | The standard blurry 4-point sampling, with some speed-up tricks.
3 | Bilinear Float | The standard blurry 4-point sampling written without tricks.
 
### 3 Point?
 I have always wondered about the "rupee" 3-Point interpolation of the N64. An optimized version can be seen in the function ReadTexel3Point().
 
 ![3 point interpolation](https://user-images.githubusercontent.com/96515734/219922319-bf7cebb7-323c-40a0-bf2a-a236b81cd49f.png)
 
 I hope to be able to describe the math behind it.
 
 When a rendered triangle is drawn larger than the size of the texture, the texture is sampled in-between the texels to magnify it more smoothly.
 
 If each of the dots of a texture are considered to be at whole number intervals, then the space between two adjacent texels is some fractional number.
 The fractional U and V coordinates vary from 0 to 0.999....
 
 Fractional_U = U - floor(U)
 
 Fractional_V = V - floor(V)
 
 When we desire to interpolate on 2 axes for the U and V coordinates it would seem we need a square. We would need to interpolate somewhere inbetween the the 4 corners.
 
(U, V) | (U+1, V)
------ | -----------
(U, V+1) | (U+1, V+1)
 
 But what if we cut the square diagnonally in half, forming two triangles that share a hypotenuse? Now there are only 3 texels to interpolate between.
 
 Independent of which half is used, two sampling coordinates will stay the same. But for the third point, it must be determined what triangle "half" an interior point is in. Which of the two possible ways the triangles are divided requires an algorith for determining which sample is chosen. This can be done by:
 1. **IF** Fractional_U + Fractional_V > 1.0  **THEN** ...Bottom Right... **ELSE** ...Top Left...
 2. **IF** Fractional_U > Fractional_V **THEN** ...Top Right... **ELSE** ...Bottom Left...
 
 Now for the blending between these three points. It is entirely possible to use the general case Barycentric math formula. It would be used to find out the 3 subdivided areas of any given point within the interior of these 3 vertexes. And then sum the ratios (weights) of the 3 areas multiplied by the color at each of the opposing vertexes arrives at the correct blended color. But to do that would require more math steps than bilinear interpolation. That is clearly not what is going on in the N64 hardware if the entire intention was cost reduction.
 
 Consider the following truths about this Barycentric triangle:
 
 1. The lengths of the legs of the outer triangle are exactly 1.0
 2. The outer triangle is always a right triangle.
 3. The area of the outer triangle is 0.5 always. (1/2 * base * height).
 4. Two of the three interior subdivided triangles can be easily decomposed into a sum of two right triangles.
 4. Two of the three interior subdivided triangles have one leg that is exactly 1.0 in length.
 5. The area of the more difficult subdivided triangle can instead be determined by subtraction from the total area 0.5
 
 The big leap in understanding is that twice the area of a right-triangle is a rectangle. If the total area of the triangle is 0.5, conceptually the area of a square is 1.0. Removing the 1/2 constant in the process of determining the areas is one less operation step. Doing that allows the fractional U or fractional V coordinate to be used directly as the area.
 
 Area = 2 * 1/2 * 1.0 * Fractional_U is just Fractional_U.

 So it ends up being 3 area multiplications per color component. For RGB, that is just 9 total multiplications.

### Texture boundaries

Nothing is really preventing the U or V texel coordinates from going outside of the range of the sampled texture. The question becomes what to do. And the answer is that it depends on what the artist wants. So it makes sense to give them the option.

1. Tile - Texture is regularly repeated. The bitwise AND function is used to keep only the lower significant bits.
2. Clamp - Texture coordinates are clamped to the min and max boundaries of the texture.

A program named *TextureWrapOptions.bas* in the Concepts folder was used to develop the Tile versus Clamp options. It draws a 2D visual as the options are changed by pressing number keys on the keyboard.

## Fog (Depth Cueing)
 Pixel Fog (Table Fog) is calculated at the end of the pixel blending process. Fog is usually intended to have objects blend into a background color with increasing distance from the viewer. Fog is just a general term, and could also represent smoke or liquid water, but I think you get the picture.
 
 It is implemented here as a gradient, where the input color values blend linearly into the fog color for Z values between the *fog_near* and *fog_far* variables. If the input Z value is closer than the *fog_near* value, the input color values are passed through unchanged. At the *fog_far* Z value and beyond, the input color is replaced with the fog color but yet the Z-Buffer is still updated and the screen pixel is still drawn.
 
 At the time of writing the code, I could not arrive at a rational reason why fog tables existed, so just the linear gradient is implemented here for simplicity sake. For different batches of triangles with different visual rendering requirements, the depths *fog_near* and *fog_far* could be adjusted accordingly.
 
 I have since learned that if the W value is used instead of the Z value for distance from the viewer, fog calculation becomes a non-linear function. So a lookup table is required to re-linearize. I have also learned that this so-called table fog was a vendor-specific implementation (3dfx). Although a few competitor chips did mostly achieve equivalency for a short while, table fog was dropped from future models in favor of vertex alpha-channel fog. The reason for that is seeing unsatisfactory or inconsistent results depending on the precision and representation (fixed or floating point) of the depth buffer. A fog table must be recalculated based on the precision of the depth buffer. Applications or games that did not, or could not, have this adjustment programmed in experienced wildly different fog results.

## Retrospective
### Introduction
 A lot of what I researched and discovered cannot be reflected in these program's source. But perhaps you would appreciate some brief random thoughts.
### Design Mindset
 It would have been tough to predict the winning companies in the PC 3D Accelerator era. Consider two approaches that could not have been more different: S3 ViRGE series versus 3dfx Voodoo series.
- S3 had an extremely well-performing and well-priced 2D accelerator Super VGA chip. To ride the wave they decided to tack on a minimalistic 3D triangle edge stepper. Their new 3D chip fit in the same PCB footprint as their earlier 2D only chip. This is the cautious retreat approach. Let's not kill our existing income base on a risky fad that might not ever materialize. Be able to roll back to safe territory when the smoke clears. Did you know the ViRGE series had hardware video playback overlay and NTSC video capture capability that also worked decently well?
- 3dfx started from scratch, and focused their limited engineering resources and transistor budget on just drawing triangles. Starting over with new hardware is usually a horrible idea because there isn't any existing software that works with it. Go big or go home. Their design had to be so blatantly obviously better in a field of competing geniuses. They won on fill rate speed, and speed indeed matters. But what to leave out, what to copy, what to skimp on, and what would waste time and be unused was largely unknowable at the time. I would say they wasted a lot of time and silicon die space on "data swizzling" (mixed big- and little endian swap conversion) features.

In the end, both companies were sold off. But it's how they're remembered and celebrated, either famously or infamously.

### Quips
 Puffery: Did you know that ViRGE stands for Virtual Reality Graphics Engine? Which is complete and utterly laughable P.T. Barnum **SUCKER** level marketing? Reading the headlined features and benefits claimed in the official datasheet is entertainment alone, compared to how it benchmarks.

 Worthless Patents: Untold time was spent on 16-bit Z-Buffers to cater to stingy manufacturers that wanted minimum viable product. *Hey, we noticed RAM is going to be very inexpensive soon, so what is the least amount of onboard RAM required?!* Is it really genius to think if you have 80-bit, 64-bit, or 32-bit floating point, that it is *possible* to have a 16-bit floating point number? But really, is that a good use of resources? By the time the Patent Lawyer's check clears, technology would have just moved on. The value of a company's patent portfolio is more about jumping out from the shadows and saying GOTCHA and hoping the competitor also made the same shortcut.

 16-bit doesn't cut it: Basically all of these first-generation 3D accelerator chips internally had 8 bits per color channel, but dithered the final colors into either RGB 555 or RGB 565, thinking we wouldn't notice. For perspective, having a healthy variety of texture color reduction and compression choices to save memory footprint leads to overall more visual variety and better graphical look. But skimping on 3 bits per channel in the final display framebuffer is just being really cheap. This isn't the first time, and it won't be the last time that dithering is used to inflate the number of bits. Well on average it's 8 bits... if you stand far enough back and squint your eyes. And yes a lot of your HDR displays are only 8-bit, twice dithered.
 
