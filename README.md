## Autodesk Inventor code samples ##

### bspline-surface ###

This example shows how to create a B-spline surface from an arbitrarily-chosen set of control points, z = 0.3 * cos(2*x^2 + 2*y^2).  
Originally this example was built to help troubleshoot the long time required to create complex splines. 
For a spline with m-by-n control points, the time seems to go as O(n^3 * m^3) in Inventor 14.

### interactive-mesh ###

This example shows how to create a static triangle mesh in the current document. (Yeah, I know the project name is misleading, but I'm too lazy to change it.)
