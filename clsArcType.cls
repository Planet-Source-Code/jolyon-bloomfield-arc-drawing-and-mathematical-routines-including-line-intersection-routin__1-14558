VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsArcDraw"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'
' This is a DLL that incorporates two important functions for graphics and maths.
' The first function is a line intersection routine that takes in two lines,
' and returns where they intersect, and if they intersect. This routine
' handles Double precision numbers, and is optimized for speed.
' The second routine draws an arc to any hDC using a width, penstyle, colour,
' and drawmode, using the API.
' Note: the scalemode of the hDC *MUST* be in pixels.
' Each arc is defined by three points. There is a data type below to describe arcs
' as used in the accompanying program. This data type can be passed into an
' alternate function for ease of use.
'
' This DLL is the result of many hours of work, mostly in maths, to calculate
' all that it does. If you use this, and please do, give me credit for my work.
'
' Otherwise, enjoy these functions, and I hope that they help you.
'
' Jolyon Bloomfield
' January 2000
' Jolyon_B@Hotmail.Com
' ICQ Uin: 11084041
'

' A standard point type in double precision, + a used variable to say that
' the point is being used
Public Type Point
  X As Double
  Y As Double
  Used As Boolean
End Type

' Declarations for each arc
Public Type Arc
  Points(1 To 3) As Point    ' The points that define the arc
  Centrepoint As Point       ' The centrepoint of the circle
  Radius As Double           ' The radius
  ClockWise As Boolean       ' True = clockwise, false = counterclockwise
  Used As Boolean            ' Is the arc being used?
  ColKey As String           ' A Collection Key used to define the arc
  Colour As Long             ' Colour of the arc
  Width As Integer           ' Width of the line on the arc
End Type

' Self explanatory
Private Const Pi = 3.14159265358979

' The results from the DrawArc and DrawArcType functions
Public Enum DrawArcResults
  NoError = 0
  COLLINEAR = 1
End Enum

' The results from the LineIntersect function
Public Enum LineIntersectResults
  DontIntersect = 1
  DoIntersect = 2
  COLLINEAR = 3
End Enum

' Declares and constants used in drawing arcs
Private Declare Function Arc Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long    ' Create a GDI graphics pen object
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long    ' Used to select brushes, pens, and clipping regions
Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long     ' Get hold of a "stock" object. I use it to get a Null Brush
Private Declare Function SetROP2 Lib "gdi32" (ByVal hdc As Long, ByVal nDrawMode As Long) As Long    ' Used to set the Raster OPeration of a window
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long    ' Delete a GDI Object
Private Const NULL_BRUSH = 5   ' Stock Object

'
' Draw an arc based upon three points (X,Y)
'
' Point1, Point2 and Point3 are those three points
' CentrePoint returns the centrepoint of the circle
' Radius returns the radius
' An hDc of 0 means "don't draw"
' Draws using the Penstyle, DrawMode, colour, and width specified
' Defaults to Solid, copyPen, Black, and 1
' Penstyle and Drawmode are the same as the VB constants
Public Function DrawArc(ByRef Point1 As Point, ByRef Point2 As Point, ByRef Point3 As Point, ByRef Centrepoint As Point, Optional ByRef Radius As Double, Optional ByRef ClockWise As Boolean, Optional ByRef hdc As Long = 0, Optional ByVal DrawMode As Integer = 13, Optional ByVal Colour As Long = 0, Optional ByVal Width As Integer = 1, Optional ByVal PenStyle As Integer = 0) As DrawArcResults

Dim TopLeft As Point
Dim BotRight As Point

DrawArc = CalcInterSect(Point1, Point2, Point3, Centrepoint) = DrawArcResults.COLLINEAR
If DrawArc = DrawArcResults.COLLINEAR Then Radius = 0: ClockWise = False: Exit Function

' Now that we have the middle, calculate the other parts needed for the arc
' Calculate the radius using distance formula
Radius = CalcRadius(Point1, Centrepoint)

' Calculate the top left and bottom right coordinates to feed into APi call "Arc"
TopLeft.X = Centrepoint.X - Radius
BotRight.X = Centrepoint.X + Radius
TopLeft.Y = Centrepoint.Y - Radius
BotRight.Y = Centrepoint.Y + Radius

' Create the pen and brush required to draw the arc
Dim RetVal As Long            ' Dummy returnvalue
Dim Pen As Long               ' Handle to a GDI Pen object
Dim Brush As Long             ' Handle to a GDI Brush object
Dim OldPen As Long            ' Handle to previous Pen object (to restore it)
Dim OldBrush As Long          ' Handle to previous brush object (to restore it)
Dim OldROP As Long            ' Value of the previous ROP

' Create a pen
Pen = CreatePen(PenStyle, Width, Colour)
' Select the pen
OldPen = SelectObject(hdc, Pen)
' Create a brush
Brush = GetStockObject(NULL_BRUSH)
' Select the brush
OldBrush = SelectObject(hdc, Brush)
' Select the ROP
OldROP = SetROP2(hdc, DrawMode)

' Calculate if the angle is clockwise or not, and draw accordingly
ClockWise = AngleWise(Point1, Point2, Point3, Centrepoint)
If ClockWise = True Then
  Arc hdc, TopLeft.X, TopLeft.Y, BotRight.X, BotRight.Y, Point3.X, Point3.Y, Point1.X, Point1.Y
Else
  Arc hdc, TopLeft.X, TopLeft.Y, BotRight.X, BotRight.Y, Point1.X, Point1.Y, Point3.X, Point3.Y
End If

'
' The Washing Up bits - VERY IMPORTANT
'
' Return the previous ROP
SetROP2 hdc, OldROP
' Return to the previous brush
SelectObject hdc, OldBrush
' Return the previous pen
SelectObject hdc, OldPen
' Delete the Brush I created
DeleteObject Brush
' Delete the Pen I created
DeleteObject Pen

End Function

' Returns the angle in radians of a line going from 0, 0, to X, Y
' by taking the arctangent, and taking into consideration the quadrant
' Returns angle from 0 to 2*Pi
Public Function ATNAngle(ByVal X As Double, ByVal Y As Double) As Double

If X = 0 Then
  If Y = 0 Then
    ATNAngle = 0
  ElseIf Y > 0 Then
    ATNAngle = Pi / 2
  Else
    ATNAngle = Pi * 1.5
  End If
ElseIf X > 0 Then
  If Y = 0 Then
    ATNAngle = 0
  ElseIf Y > 0 Then
    ATNAngle = Atn(Y / X)
  Else
    ATNAngle = 2 * Pi + Atn(Y / X)
  End If
Else
  ' X < 0
  If Y = 0 Then
    ATNAngle = 0
  ElseIf Y > 0 Then
    ATNAngle = Pi + Atn(Y / X)
  Else
    ATNAngle = Pi + Atn(Y / X)
  End If
End If

End Function

'
' Returns the coordinates at which two lines intersect
' If they don't intersect, the values are set to 0, and the return value set
' accordingly.
' However, if the CalcRegardless value is set to true, it will return the values
' regardless of whether they intersect or not, unless they are collinear.
'
' Note: This is quite heavy maths; I used a book to assist me in creating this
' function
'

Public Function LineIntersect(ByRef Line1P1 As Point, ByRef Line1P2 As Point, ByRef Line2P1 As Point, ByRef Line2P2 As Point, ByRef ReturnCoords As Point, Optional CalcRegardless As Boolean = False) As LineIntersectResults

Dim Ax As Double, Bx As Double, Ay As Double, By As Double, Cx As Double, Cy As Double, D As Double, E As Double, F As Double, Temp As Double, Num As Double, Offset As Double
Dim x1lo As Double, x1hi As Double, y1lo As Double, y1hi As Double

Dim X1 As Double
Dim X2 As Double
Dim X3 As Double
Dim X4 As Double
Dim X As Double
Dim Y1 As Double
Dim Y2 As Double
Dim Y3 As Double
Dim Y4 As Double
Dim Y As Double

X1 = Line1P1.X
X2 = Line1P2.X
X3 = Line2P1.X
X4 = Line2P2.X
Y1 = Line1P1.Y
Y2 = Line1P2.Y
Y3 = Line2P1.Y
Y4 = Line2P2.Y

' Check to see whether the lines fall within the bounding box of each other
Ax = X2 - X1
Bx = X3 - X4

If Ax < 0 Then
  x1lo = X2
  x1hi = X1
Else
  x1lo = X1
  x1hi = X2
End If
If Bx > 0 Then
  If x1hi < X4 Or X3 < x1lo Then
    LineIntersect = LineIntersectResults.DontIntersect
    ReturnCoords.X = 0: ReturnCoords.Y = 0
    If CalcRegardless = False Then Exit Function
  End If
Else
  If x1hi < X3 Or X4 < x1lo Then
    LineIntersect = LineIntersectResults.DontIntersect
    If CalcRegardless = False Then ReturnCoords.X = 0: ReturnCoords.Y = 0: Exit Function
  End If
End If

Ay = Y2 - Y1
By = Y3 - Y4

If Ay < 0 Then
  y1lo = Y2
  y1hi = Y1
Else
  y1lo = Y1
  y1hi = Y2
End If

If By > 0 Then
  If y1hi < Y4 Or Y3 < y1lo Then LineIntersect = LineIntersectResults.DontIntersect: If CalcRegardless = False Then ReturnCoords.X = 0: ReturnCoords.Y = 0: Exit Function
Else
  If y1hi < Y3 Or Y4 < y1lo Then LineIntersect = LineIntersectResults.DontIntersect: If CalcRegardless = False Then ReturnCoords.X = 0: ReturnCoords.Y = 0: Exit Function
End If

Cx = X1 - X3
Cy = Y1 - Y3
D = By * Cx - Bx * Cy
F = Ay * Bx - Ax * By
If F = 0 Then LineIntersect = LineIntersectResults.COLLINEAR:  ReturnCoords.X = 0: ReturnCoords.Y = 0: Exit Function
If F > 0 Then
  If D < 0 Or D > F Then LineIntersect = LineIntersectResults.DontIntersect: If CalcRegardless = False Then ReturnCoords.X = 0: ReturnCoords.Y = 0: Exit Function
Else
  If D > 0 Or D < F Then LineIntersect = LineIntersectResults.DontIntersect: If CalcRegardless = False Then ReturnCoords.X = 0: ReturnCoords.Y = 0: Exit Function
End If

E = Ax * Cy - Ay * Cx
If F > 0 Then
  If E < 0 Or E > F Then LineIntersect = LineIntersectResults.DontIntersect: If CalcRegardless = False Then ReturnCoords.X = 0: ReturnCoords.Y = 0: Exit Function
Else
  If E > 0 Or E < F Then LineIntersect = LineIntersectResults.DontIntersect: If CalcRegardless = False Then ReturnCoords.X = 0: ReturnCoords.Y = 0: Exit Function
End If

' Compute intersection coordinates
Num = D * Ax
Offset = IIf(Same_Signs(Num, F), F / 2, -F / 2)
X = X1 + (Num + Offset) / F

Num = D * Ay
Offset = IIf(Same_Signs(Num, F), F / 2, -F / 2)
Y = Y1 + (Num + Offset) / F

ReturnCoords.X = X
ReturnCoords.Y = Y

If LineIntersect = 0 Then LineIntersect = LineIntersectResults.DoIntersect

End Function

' Returns whether or not A and B are the same sign (Neg or pos)
Private Function Same_Signs(ByVal A As Double, ByVal B As Double) As Boolean
If (A > 0 And B > 0) Then
  ' both pos
  Same_Signs = True
ElseIf (A < 0 And B < 0) Then
  ' both neg
  Same_Signs = True
Else
  ' one pos, one neg
  Same_Signs = False
End If
End Function

'
' Draw an arc based upon the structure "Arc" - uses the procedure
' "DrawArc"
'
Public Function DrawArcType(ByRef Arc As Arc, Optional ByVal hdc As Long = 0, Optional ByVal DrawMode As Integer = 13, Optional ByVal PenStyle As Integer = 0) As DrawArcResults
DrawArcType = DrawArc(Arc.Points(1), Arc.Points(2), Arc.Points(3), Arc.Centrepoint, Arc.Radius, Arc.ClockWise, hdc, DrawMode, Arc.Colour, Arc.Width, PenStyle)
End Function

'
' Return the radius of an arc; also place it into the data type
' MAKE SURE THAT THE 3 POINTS ARE SET BEFORE DOING THIS
' ALSO MAKE SURE TO CALCULATE THE CENTREPOINT FIRST, TOO!!!
'
Public Function ArcRadius(ByRef Arc As Arc) As Double
ArcRadius = CalcRadius(Arc.Points(1), Arc.Centrepoint)
Arc.Radius = ArcRadius
End Function
Public Function CalcRadius(ByRef Point1 As Point, ByRef Centrepoint As Point) As Double
CalcRadius = Sqr((Centrepoint.X - Point1.X) ^ 2 + (Centrepoint.Y - Point1.Y) ^ 2)
End Function

'
' Calculate the intersection point of three lines, also
' add this to the Arc type if requested
'
Public Function CalcInterSectArc(ByRef Arc As Arc) As DrawArcResults
CalcInterSectArc = CalcInterSect(Arc.Points(1), Arc.Points(2), Arc.Points(3), Arc.Centrepoint)
End Function
Public Function CalcInterSect(ByRef Point1 As Point, ByRef Point2 As Point, ByRef Point3 As Point, ByRef Centrepoint As Point) As DrawArcResults
Dim I As Integer
Dim Side1(1 To 2) As Point
Dim Side2(1 To 2) As Point

' First step is to find the middle
' First, get the halfway distance between two points
Side1(1).X = (Point1.X + Point2.X) / 2
Side1(1).Y = (Point1.Y + Point2.Y) / 2
' Then, calculate to create a line perpendicular to the one above
Side1(2).X = Side1(1).X + (Point1.Y - Point2.Y)
Side1(2).Y = Side1(1).Y - (Point1.X - Point2.X)

' Now, do it again for the next side
Side2(1).X = (Point3.X + Point2.X) / 2
Side2(1).Y = (Point3.Y + Point2.Y) / 2
Side2(2).X = Side2(1).X + (Point3.Y - Point2.Y)
Side2(2).Y = Side2(1).Y - (Point3.X - Point2.X)

' Next, find the point at which these two lines cross - the centrepoint
If LineIntersect(Side1(1), Side1(2), Side2(1), Side2(2), Centrepoint, True) = LineIntersectResults.COLLINEAR Then
  ' The three points are collinear (In a straight line), so an arc cannot be created
  CalcInterSect = DrawArcResults.COLLINEAR
  Centrepoint.X = 0
  Centrepoint.Y = 0
End If

End Function

'
' Calculates the angles at which the points lie
' MUST have CalcIntersect done BEFORE RUNNING THIS!
'
Public Function AngleWiseArc(ByRef Arc As Arc) As Boolean
AngleWiseArc = AngleWise(Arc.Points(1), Arc.Points(2), Arc.Points(3), Arc.Centrepoint)
Arc.ClockWise = AngleWiseArc
End Function
Public Function AngleWise(ByRef Point1 As Point, ByRef Point2 As Point, ByRef Point3 As Point, ByRef Centrepoint As Point) As Boolean
Dim Angle1 As Double
Dim Angle2 As Double
Dim Angle3 As Double

' Calculate all three angles
Angle1 = ATNAngle(Point1.X - Centrepoint.X, Centrepoint.Y - Point1.Y)
Angle2 = ATNAngle(Point2.X - Centrepoint.X, Centrepoint.Y - Point2.Y)
Angle3 = ATNAngle(Point3.X - Centrepoint.X, Centrepoint.Y - Point3.Y)

AngleWise = CalcAngleWay(Angle1, Angle2, Angle3)

End Function

'
' Returns if an angle is stored inbetween two other angles.
' In the case of the Arc encompassing, the first and third points are used
' to see if another angle lies in the arc
'
Public Function AngleEncompassedArc(ByRef Arc As Arc, ByVal AngleSeek As Double) As Boolean
AngleEncompassedArc = AngleEncompassed(ATNAngle(Arc.Points(1).X - Arc.Centrepoint.X, Arc.Centrepoint.Y - Arc.Points(1).Y), ATNAngle(Arc.Points(3).X - Arc.Centrepoint.X, Arc.Centrepoint.Y - Arc.Points(3).Y), AngleSeek, Arc.ClockWise)
End Function
Public Function AngleEncompassed(ByVal Angle1 As Double, ByVal Angle2 As Double, ByVal AngleSeek As Double, ByVal ClockWise As Boolean) As Boolean
AngleEncompassed = (CalcAngleWay(Angle1, AngleSeek, Angle2) = ClockWise)
'Debug.Assert AngleEncompassed = False
End Function

' Assists in AngleEncompassed and AngleWise
Private Function CalcAngleWay(ByVal Angle1 As Double, ByVal Angle2 As Double, ByVal Angle3 As Double) As Boolean
If Angle1 > Angle2 And (Angle2 > Angle3 Or Angle3 > Angle1) Then
  CalcAngleWay = True
ElseIf Angle1 < Angle2 And Angle3 > Angle1 And Angle3 < Angle2 Then
  CalcAngleWay = True
Else
  CalcAngleWay = False
End If
End Function
