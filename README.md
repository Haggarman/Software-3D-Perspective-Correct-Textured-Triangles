# Language
 This BASIC program is written for the QB64 compiler.
 https://github.com/QB64-Phoenix-Edition/
#Description
 How were 3D triangles drawn on the first graphics accelerators? This was my deep dive into understanding the software algorithms involved in drawing triangles line by line.
#Approach
 For the sake of understanding, floating point numbers are used. On early accelerators there were many different fixed-point number combinations for the sake of speed and reduced transistor count. But all that shifting and conversion hinders understanding.
 
#Texture Magnification
 I have always wondered about the "rupee" 3-Point interpolation of the N64. It is there in the function ReadTexel3Point( ). I hope to be able to describe the math behind it. The big leap in understanding is that twice the area of a right-triangle is a rectangle, and that greatly simplifies the calculation to just a multiplication.
 