<div align="center">

## Arc Drawing and Mathematical Routines \- Including Line Intersection Routines \- \*Updated\!\*


</div>

### Description

This submission is split into 2 - A DLL, which holds all routines for handling arcs, and a program, which demonstrates how to use the DLL.

The DLL is basically a maths DLL, exposing a Line Intersection Routine, and an Arc Drawing routine. Each arc is defined by a series of three points - three points along the line that is the arc. The arc is drawn to any hDC by API, using specified colours, widths, penstyles, and drawmodes. The DLL also contains data type defining arcs, and routines to suit, to make using this DLL as easy as possible.

The program is a mimick of a CAD package, except it only uses arcs. It allows creation of arcs by clicking three points on a picturebox, then the arc, and the centrepoint of its circle, are drawn. From here, the points may be moved to update the arc, by dragging them and dropping them. During a dragdrop operation, the arc can be seen to be moving by using XOr drawmodes.

The update includes the following changes:

Program:

The Arcs may now be moved by dragging the arc where there is no control point.

An Arc must be selected before its control points can be moved.

Added Arc Selection Routine to Point Selection Routine.

About and Help boxes added.

Arcs can now be rotated around the centrepoint by rightclicking and dragging.

ZOrder for arcs also added, + buttons to manipulate.

DLL:

Updated routines and handling of data types inside DLL.

Numerous routines added to assist external calculations of angles, distances, radii, etc.

Overall, this has been a wonderful experience with mathematics for me. I strongly suggest that you download it, even if it is just to see what VB is really capable of, when the maths is applied. My 486 is fast enough to draw the arcs in real time, so I believe that any computer should do justice to the wonder of maths and graphics combined.

Hope you enjoy it!

Jolyon Bloomfield
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2001-01-20 19:21:06
**By**             |[Jolyon Bloomfield](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jolyon-bloomfield.md)
**Level**          |Advanced
**User Rating**    |4.8 (38 globes from 8 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Math/ Dates](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/math-dates__1-37.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[CODE\_UPLOAD139521202001\.zip](https://github.com/Planet-Source-Code/jolyon-bloomfield-arc-drawing-and-mathematical-routines-including-line-intersection-routin__1-14558/archive/master.zip)








