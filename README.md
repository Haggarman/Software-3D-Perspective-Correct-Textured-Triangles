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
 
## Approach
 QBASIC was used because of the very quick edit-compile-test iterations. I wanted this exploration to be fun.
 
 A verbose style of naming variables and keeping ideas separated per source code line is used.
 
 Floating point numbers are used for the sake of understanding. On early 3D accelerators there were many different fixed-point number combinations for the sake of speed and reduced transistor count. But all that bit shifting hinders understanding.
 
 I believe it would be much easier to translate this BASIC code to C-lang or Python, than it was for me to catch onto the nuances from other's example C++ code that was using STL std::list, templates and pointer tricks.
 
 No dropping to assembly or using pokes!
 
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

