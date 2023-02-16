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
 I have always wondered about the "rupee" 3-Point interpolation of the N64. It is there in the function ReadTexel3Point().
 
 I hope to be able to describe the math behind it. The idea is that of course 3 points determine a triangle (or a circle, or a rectangle). But we're talking about a triangle.
 
 The U and V coordinates vary from 0 to 0.999.... Fractional_U = U - floor(U)
 
 When we desire to interpolate on 2 axes for the U and V coordinates it would seem we need a square. We would need to interpolate between the colors at the 4 corners. But what if we cut the square in half, forming two triangles that share a hypotenuse?
 
 Independent of which half is used, two sampling coordinates stay the same. But for the third point, it must be determined what triangle "half" an interior point is in. This can be done by:
 1. IF U+V > 1.0  THEN ... ELSE ...
 2. IF U > V THEN ... ELSE ...
 
 Which of the two is chosen affects which of the two possible ways the triangles are divided. Top-Left / Bottom-Right. Bottom-Left \ Top-Right.
 
 It is entirely possible to use generic Barycentric math to find out the ratios of areas of any given point within the triangle and the 3 vertexes. And then use the ratio of the areas divided by the total area to arrive at the correct color component. But consider the three following assumptions:
 
 1. The triangle is always a right triangle.
 2. The lengths of the legs of the triangle are exactly 1.0
 3. The area of the triangle is 0.5. (1/2 * base * height).
 
 The big leap in understanding is that twice the area of a right-triangle is a rectangle. If the total area of the triangle is 0.5, conceptually the area of a square is 1.0. That greatly simplifies the calculation to just a multiplication.

