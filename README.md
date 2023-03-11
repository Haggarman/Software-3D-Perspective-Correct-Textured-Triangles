## Language
 This collection of BASIC programs is written for the QB64 compiler.
 https://github.com/QB64-Phoenix-Edition/
## Description
 How were 3D triangles drawn by the first PC graphics accelerators in 1997? This was my deep dive into understanding the software algorithms involved in drawing triangles line by line. This is a software simulation only approach.
 
 For extra fun there are blending operations like Fog, Vertex Color, and Directional/Ambient lighting.
## Screenshots
 ![Textured Cube Plains](https://user-images.githubusercontent.com/96515734/219141104-2c77c587-f4d6-4f52-baaa-d302638e0d97.PNG)
 ![Vertex Color Cube](https://user-images.githubusercontent.com/96515734/219141230-6d9a9d36-ec02-4c5a-94eb-a01773a02b5b.PNG)
 ![Dither Color Cube](https://user-images.githubusercontent.com/96515734/219270641-b560848f-5ef0-429c-9c9c-6191e8ac88a9.png)
 ![Skybox Long Tube](https://user-images.githubusercontent.com/96515734/222982848-81e9a2b5-2fa4-4dbf-af8a-e3ee4164b683.png)
 
## Approach
 QBASIC was used because of the very quick edit-compile-test iterations. I wanted this exploration to be fun.
 
 A verbose style of naming variables and keeping ideas separated per source code line is used.
 
 Floating point numbers are used for the sake of understanding. On early 3D accelerators there were many different fixed-point number combinations for the sake of speed and reduced transistor count. But all that bit shifting hinders understanding.
 
 I believe it would be much easier to translate this BASIC code to C-lang or Python, than it was for me to catch onto the nuances from other's example C++ code that was using STL std::list, templates and pointer tricks.
 
 No dropping to assembly or using pokes!
## Triangles
### Vertex
 The triangles are specified by vertexes A, B, and C. They are sorted by the triangle drawing subroutine so that A is always on top and C is always on the bottom. That still leaves two categories where the knee at B faces left or right. The triangle drawing subroutine also adjusts for this so that pixels are drawn from left to right.
 ![TrianglesABC](https://user-images.githubusercontent.com/96515734/220204499-62aaed3c-f1fe-4c07-9c64-1c61564219e7.PNG)
### DDA
 The DDA (Digital Difference Analyzer) algorithm is used to simultaneously step on whole number Y increments from point A to point C on the major edge, and from point A to point B on the minor edge. The start value of Y at point A is pre-stepped ahead to the next highest integer pixel row using the ceiling function.
 
 To ensure that the sampling is visually correct, the X major, X minor, and vertex attributes (U, V, R, G, B, etc.) are also pre-stepped forward by the same amount. This prestep of Y also factors in the clipping window so that the DDA accumulators are correctly advanced to the top row of the clipping region.
### Knee B
 When the Minor Edge DDAs reach vertexBy, the start values and steps are recalculated to be from point B to point C. Note that this case also handles a flat-topped triangle where vertexBy = vertexAy.
 
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

The sign (positive or negative) of the triangle's **surface normal** as compared to a normalized ray extending out from the viewer (using the dot product) can determine which side of the triangle is facing the viewer. The Z value of this surface normal can be used if re-calculated after rotation and translation but before projection, because the view normal is usually (0, 0, 1). If the triangle were to be viewed perfectly edge-on to have a value of 0, it is also invisible because it is infinitely thin. So not drawing the triangle if this value is less than or equal to 0.0 accomplishes backface culling.

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
3. Decal - Texture simply is not drawn if coordinates are outside the boundaries of the texture.

A program named *TextureWrapOptions.bas* in the Concepts folder was used to develop the Tile versus Clamp options. It draws a 2D visual as the options are changed by pressing number keys on the keyboard.

## Fog (Depth Cueing)
 Fog is calculated at the end of the pixel blending process. Fog is usually intended to have objects blend into a background color with increasing distance from the viewer.
 
 It is implemented here as a gradient, where the input color values blend linearly into the fog color for Z values between the fog_near and fog_far variables. If the input Z value is closer than the fog_near value, the input color values are passed through unchanged. At the fog_far Z value and beyond, the input color is replaced with the fog color but yet the Z-Buffer is still updated and the screen pixel is still drawn.
 
 I could not arrive at a rational reason why fog tables existed, so just the gradient is implemented here for simplicity sake. For different batches of triangles with different visual rendering requirements, the depths fog_near and fog_far could be adjusted accordingly.
