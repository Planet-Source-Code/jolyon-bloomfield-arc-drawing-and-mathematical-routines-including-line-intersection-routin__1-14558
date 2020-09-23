VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmArc 
   AutoRedraw      =   -1  'True
   Caption         =   "Vector Arcs"
   ClientHeight    =   5565
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9075
   Icon            =   "frmArc.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   371
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   605
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7200
      Top             =   5040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdZOrder 
      Caption         =   "Send to Top"
      Height          =   495
      Index           =   3
      Left            =   5340
      TabIndex        =   14
      Top             =   5040
      Width           =   1275
   End
   Begin VB.CommandButton cmdZOrder 
      Caption         =   "Move Up"
      Height          =   495
      Index           =   2
      Left            =   4020
      TabIndex        =   13
      Top             =   5040
      Width           =   1275
   End
   Begin VB.CommandButton cmdZOrder 
      Caption         =   "Move Down"
      Height          =   495
      Index           =   1
      Left            =   2700
      TabIndex        =   12
      Top             =   5040
      Width           =   1275
   End
   Begin VB.CommandButton cmdZOrder 
      Caption         =   "Send to Bottom"
      Height          =   495
      Index           =   0
      Left            =   1380
      TabIndex        =   11
      Top             =   5040
      Width           =   1275
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Help"
      Height          =   495
      Left            =   60
      TabIndex        =   9
      Top             =   2760
      Width           =   1275
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "About"
      Height          =   495
      Left            =   60
      TabIndex        =   8
      Top             =   3300
      Width           =   1275
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Show Centrepoints"
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   4500
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Show Control Points"
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   3840
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.PictureBox picArc 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   4935
      Left            =   1380
      ScaleHeight     =   325
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   505
      TabIndex        =   6
      Top             =   60
      Width           =   7635
   End
   Begin VB.CommandButton cmdWidth 
      Caption         =   "Width"
      Height          =   495
      Left            =   60
      TabIndex        =   4
      Top             =   2220
      Width           =   1275
   End
   Begin VB.CommandButton cmdColour 
      Caption         =   "Colour"
      Height          =   495
      Left            =   60
      TabIndex        =   3
      Top             =   1680
      Width           =   1275
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "Delete Arc"
      Height          =   495
      Left            =   60
      TabIndex        =   1
      Top             =   600
      Width           =   1275
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New Arc"
      Height          =   495
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   1275
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear All Arcs"
      Height          =   495
      Left            =   60
      TabIndex        =   2
      Top             =   1140
      Width           =   1275
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ZOrder:"
      Height          =   195
      Left            =   720
      TabIndex        =   10
      Top             =   5160
      Width           =   540
   End
End
Attribute VB_Name = "frmArc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
' This Project is an example of how to use the arc drawing capabilities of the
' included DLL/Project. It is a simple Vector drawing program that allows
' positioning of arcs using three points to define each arc, then drawing
' it with width and colour. It also allows movement of the points by dragging them-
' a rather nice effect. Enjoy this program.
' If you use this, please give me credit for the code that I have laborously
' created.
'
' Jolyon Bloomfield
' Jolyon_B@Hotmail.Com
' ICQ Uin: 11084041
'

Private Const Pi = 3.14159265358979

' Used for creating points
Private Upto As Integer
Private Creating As Boolean

' The array that stores all of the arcs
Private Arcs() As Arc

' Used for selecting points
Private CurArcKey As String
Private Const Tolerance = 3

' Used for moving points
Private Moving As Boolean
Private MovePoint As Point     'Used to store the index and the point number
Private MoveStart As Point

Private Sub Check1_Click()
' Redraw all the arcs, to update the way that the dots are shown
ReDraw
End Sub

Private Sub Check2_Click()
' Redraw all the arcs, to update the way that the radii are shown
ReDraw
End Sub

Private Sub cmdAbout_Click()
MsgBox "This program written entirely by Jolyon Bloomfield January 2001. If you use any portion of this program, please give credit to Jolyon." & vbCrLf & "Bugs, comments, requests, queries, etc, please contact Jolyon by E-mail: Jolyon_B@Hotmail.Com", vbInformation, "About ArcDraw Dll and example program"
End Sub

Private Sub cmdClear_Click()
' Clear everything
picArc.Cls
Upto = 0
Creating = False
picArc.MousePointer = 0
CurArcKey = ""
ReDim Arcs(0 To 0) As Arc
cmdNew.Enabled = True
End Sub

Private Sub cmdColour_Click()
' Choose and change the colour of an arc, then redraw
If CurArcKey = "" Then MsgBox "Please select an arc.", vbInformation, "No Arc Selected": Exit Sub
CommonDialog1.Color = Arcs(GetIndex(CurArcKey)).Colour
CommonDialog1.CancelError = True
On Error Resume Next
CommonDialog1.ShowColor
If Err Then Exit Sub
On Error GoTo 0
If CommonDialog1.Color = Arcs(GetIndex(CurArcKey)).Colour Then Exit Sub
picArc.Cls
Arcs(GetIndex(CurArcKey)).Colour = CommonDialog1.Color
ReDraw
End Sub

Private Sub cmdDel_Click()
' Remove all entries for an arc
If CurArcKey = "" Then MsgBox "Please select an arc.", vbInformation, "No Arc Selected": Exit Sub
picArc.Cls
Upto = 0
Creating = False
picArc.MousePointer = 0
Arcs(GetIndex(CurArcKey)).Used = False
CurArcKey = ""
cmdNew.Enabled = True
ReDraw
End Sub

Private Sub cmdHelp_Click()
MsgBox "How to use:" & vbCrLf & "Step 1. Click on ""New Arc""" & vbCrLf & "Step 2. Click 3 points on the white area to create an arc." & vbCrLf & "Repeat to create more arcs." & vbCrLf _
     & "To select an arc, click on it. When an arc is selected, you may change its colour and width, or delete it." & vbCrLf & "To delete all arcs, click on ""Clear All Arcs""." & vbCrLf _
     & "To hide display of Centrepoints or Control Points, uncheck the appropriate checkboxes. Note that the currently selected arc will have these displayed anyway." & vbCrLf & _
     "An arc may be moved by selecting it and dragging it. Control Points may be moved individually by clicking on them and dragging them to their new position." & vbCrLf & _
     "An arc may be rotated after selecting it by rightclicking and dragging it around its centrepoint." & vbCrLf _
     & "The ZOrder Buttons are used to control the layering of the arcs, on top of each other - the positioning system should be fairly easy to understand, and its effect obvious, when arcs are made both wide and different colours." _
     & "This program is made to demonstrate the ArcDraw DLL. I hope you enjoy it. Jolyon Bloomfield January 2000   Jolyon_B@Hotmail.Com"
End Sub

Private Sub cmdNew_Click()
' Create a new arc, by adding an entry
Deselect
Creating = True
picArc.MousePointer = 2
Upto = 0
cmdNew.Enabled = False
Dim NewArc As Arc
NewArc.ColKey = GetNewKey
CurArcKey = NewArc.ColKey
NewArc.Colour = 0
NewArc.Width = 1
Arcs(GetNewIndex) = NewArc
End Sub

Private Sub cmdWidth_Click()
' Sets the width of an arc, then redraws
If CurArcKey = "" Then MsgBox "Please select an arc.", vbInformation, "No Arc Selected": Exit Sub
Dim Temp As Integer
Temp = frmGetWidth.GetNewWidth(Arcs(GetIndex(CurArcKey)).Width)
If Temp = Arcs(GetIndex(CurArcKey)).Width Then Exit Sub
picArc.Cls
Arcs(GetIndex(CurArcKey)).Width = Temp
ReDraw
End Sub

Private Sub cmdZOrder_Click(Index As Integer)
If CurArcKey = "" Then MsgBox "Please select an arc.", vbInformation, "No Arc Selected": Exit Sub
' Start by compressing the array - removing unused elements
CompressArray
If UBound(Arcs) < 2 Then MsgBox "There needs to be more arcs to use this function.", vbInformation, "Needs more arcs": Exit Sub

Dim I As Integer
Dim Temp As Arc

Select Case Index
  Case Is = 0 'Send to Bottom
    Temp = Arcs(GetIndex(CurArcKey))
    For I = GetIndex(CurArcKey) To UBound(Arcs) - 1
      Arcs(I) = Arcs(I + 1)
    Next I
    Arcs(UBound(Arcs)) = Temp
  Case Is = 3 ' Send to Top
    Temp = Arcs(GetIndex(CurArcKey))
    For I = GetIndex(CurArcKey) To 2 Step -1
      Arcs(I) = Arcs(I - 1)
    Next I
    Arcs(1) = Temp
  Case Is = 1 ' Move down
    I = GetIndex(CurArcKey)
    Temp = Arcs(I)
    Arcs(I) = Arcs(I + 1)
    Arcs(I + 1) = Temp
  Case Is = 2 ' Move up
    I = GetIndex(CurArcKey)
    Temp = Arcs(I)
    Arcs(I) = Arcs(I - 1)
    Arcs(I - 1) = Temp
End Select

ReDraw

End Sub

Private Sub Form_Load()
' Dimension the Arcs array
ReDim Arcs(0 To 0) As Arc
End Sub

Private Sub picArc_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' If creating an arc, then get out!
If Creating = True Then Exit Sub

Dim Out As Point
Dim I As Integer

If Button = vbLeftButton Then
  ' Enter Point Selection Routine
  Out = FindPoint(X, Y)
  If Out.X = 0 Then
    ' No points selected
  Else
    If Arcs(Out.X).ColKey <> CurArcKey Then
      ' New arc selected
      ' Deselect the old one
      Deselect
      ' Select the new one
      SelectArc Arcs(Out.X)
      ' Set moving to false...
      Moving = False
    End If
    ' Prepare to move the point
    Moving = True
    MovePoint.X = Out.X
    MovePoint.Y = Out.Y
    Arcs(0).Used = True
    Arcs(0) = Arcs(Out.X)
    ' Get out so that the moving can take place
    Exit Sub
  End If
  ' Enter Arc Selection Routine
  I = FindArc(X, Y)
  If I = 0 Then
    ' No points selected, deselect, and set moving to false
    Moving = False
    Deselect
  Else
    If Arcs(I).ColKey <> CurArcKey Then
      ' New arc selected
      ' Deselect the old arc
      Deselect
      ' Select the new arc
      SelectArc Arcs(I)
      ' Set moving to false...
      Moving = False
    End If
    ' Prepare to move the arc
    Moving = True
    MovePoint.X = I
    MovePoint.Y = 0       ' Tag to say to move the entire arc
    MoveStart.X = X
    MoveStart.Y = Y
    Arcs(0).Used = True
    Arcs(0) = Arcs(I)
    ' Get out so that the moving can take place
    Exit Sub
  End If
ElseIf Button = vbRightButton Then
  I = FindArc(X, Y)
  If I = 0 Then
    ' No points selected, deselect, and set moving to false
    Moving = False
    Deselect
  Else
    If Arcs(I).ColKey <> CurArcKey Then
      ' New arc selected
      Moving = False
      ' Get out - don't want to rotate
      Exit Sub
    End If
    ' Prepare to move the arc
    Moving = True
    MovePoint.X = I
    MovePoint.Y = -1      ' Tag to say to rotate the arc
    MoveStart.X = X
    MoveStart.Y = Y
    Arcs(0).Used = True
    Arcs(0) = Arcs(I)
    ' Get out so that the moving can take place
    Exit Sub
  End If
End If

End Sub

Private Sub picArc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim I As Integer
Dim RotateAngle As Double
Dim Temp As Double
Dim Dist As Double
If Button = vbLeftButton And Moving = True Then
  If MovePoint.Y > 0 Then
    ' If points are moving
    If Arcs(0).Used = True Then
      'If there's already a "moving" arc drawn, undraw it.
      ' The drawmode is set to 6 - PS_Invert
      DrawArcType Arcs(0), picArc.hDC, 6
    End If
    Arcs(0).Used = True
    Arcs(0) = Arcs(GetIndex(CurArcKey))
    ' Update the coordinates
    Arcs(0).Points(MovePoint.Y).X = X
    Arcs(0).Points(MovePoint.Y).Y = Y
    ' Draw the new "moving" arc - a preview of what the new arc will look like
    DrawArcType Arcs(0), picArc.hDC, 6
    ' Refresh the picturebox
    picArc.Refresh
  Else      ' Move the entire Arc
    If Arcs(0).Used = True Then
      'If there's already a "moving" arc drawn, undraw it.
      ' The drawmode is set to 6 - PS_Invert
      DrawArcType Arcs(0), picArc.hDC, 6
    End If
    Arcs(0).Used = True
    Arcs(0) = Arcs(GetIndex(CurArcKey))
    ' Update the coordinates
    For I = 1 To 3
      Arcs(0).Points(I).X = Arcs(0).Points(I).X + X - MoveStart.X
      Arcs(0).Points(I).Y = Arcs(0).Points(I).Y + Y - MoveStart.Y
    Next I
    ' Draw the new "moving" arc - a preview of what the new arc will look like
    DrawArcType Arcs(0), picArc.hDC, 6
    ' Refresh the picturebox
    picArc.Refresh
  End If
ElseIf Button = vbRightButton Then
  If MovePoint.Y = -1 Then
    If Arcs(0).Used = True Then
      'If there's already a "moving" arc drawn, undraw it.
      ' The drawmode is set to 6 - PS_Invert
      DrawArcType Arcs(0), picArc.hDC, 6
    End If
    Arcs(0).Used = True
    Arcs(0) = Arcs(GetIndex(CurArcKey))
    ' Rotate all three coordinates around the centrepoint
    ' First, calculate the angle that the mouse has moved by
    CalcIntersectArc Arcs(0)
    RotateAngle = ATNAngle(X - Arcs(0).centrepoint.X, Y - Arcs(0).centrepoint.Y) _
                  - ATNAngle(MoveStart.X - Arcs(0).centrepoint.X, MoveStart.Y - _
                  Arcs(0).centrepoint.Y)
    ' Now update all three points by this angle
    For I = 1 To 3
      ' Find the angle to be rotated by
      Temp = ATNAngle(Arcs(0).Points(I).X - Arcs(0).centrepoint.X, Arcs(0).Points(I).Y - Arcs(0).centrepoint.Y)
      Temp = Temp + RotateAngle
      ' Find the distance
      Dist = Distance(Arcs(0).centrepoint, Arcs(0).Points(I))
      ' Calculate the new point, putting it straight into the data type
      BAndD Arcs(0).centrepoint.X, Arcs(0).centrepoint.Y, Temp, Dist, Arcs(0).Points(I)
    Next I
    ' Draw the new "moving" arc - a preview of what the new arc will look like
    DrawArcType Arcs(0), picArc.hDC, 6
    ' Refresh the picturebox
    picArc.Refresh
  End If
End If

End Sub

Private Sub PicArc_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim I As Integer
Dim RotateAngle As Double
Dim Temp As Double
Dim Dist As Double
If Creating = True Then
  ' If an arc is being created (Points being placed)
  ' then add the XY point information, and draw that point.
  If Button = vbLeftButton Then
    Select Case Upto
      Case Is = 0, 1, 2
        Upto = Upto + 1
        Arcs(GetIndex(CurArcKey)).Points(Upto).X = X
        Arcs(GetIndex(CurArcKey)).Points(Upto).Y = Y
        Arcs(GetIndex(CurArcKey)).Points(Upto).Used = True
        Arcs(GetIndex(CurArcKey)).Used = True
    End Select
    ' When finished, update everything
    If Upto = 3 Then Creating = False: picArc.MousePointer = 0: Upto = 0: cmdNew.Enabled = True
    PutArc CurArcKey
  End If
  Exit Sub
End If

If Moving = True Then
  If MovePoint.Y > 0 Then
    ' If things are moving, then move the points, update data, set flags, and redraw
    Arcs(MovePoint.X).Points(MovePoint.Y).X = X
    Arcs(MovePoint.X).Points(MovePoint.Y).Y = Y
    ReDraw
    Moving = False
    Arcs(0).Used = False
  ElseIf MovePoint.Y = 0 Then
    ' If the whole thing has moved, update all of the points, and redraw, after updating the flags
    For I = 1 To 3
      Arcs(MovePoint.X).Points(I).X = Arcs(MovePoint.X).Points(I).X + X - MoveStart.X
      Arcs(MovePoint.X).Points(I).Y = Arcs(MovePoint.X).Points(I).Y + Y - MoveStart.Y
    Next I
    ReDraw
    Moving = False
    Arcs(0).Used = False
  ElseIf MovePoint.Y = -1 Then
    ' Rotation
    ' Rotate all three coordinates around the centrepoint
    ' First, calculate the angle that the mouse has moved by
    CalcIntersectArc Arcs(MovePoint.X)
    RotateAngle = ATNAngle(X - Arcs(MovePoint.X).centrepoint.X, Y - Arcs(MovePoint.X).centrepoint.Y) _
                  - ATNAngle(MoveStart.X - Arcs(MovePoint.X).centrepoint.X, MoveStart.Y - _
                  Arcs(MovePoint.X).centrepoint.Y)
    ' Now update all three points by this angle
    For I = 1 To 3
      ' Find the angle to be rotated by
      Temp = ATNAngle(Arcs(MovePoint.X).Points(I).X - Arcs(MovePoint.X).centrepoint.X, Arcs(MovePoint.X).Points(I).Y - Arcs(MovePoint.X).centrepoint.Y)
      Temp = Temp + RotateAngle
      ' Find the distance
      Dist = Distance(Arcs(MovePoint.X).centrepoint, Arcs(MovePoint.X).Points(I))
      ' Calculate the new point, putting it straight into the data type
      BAndD Arcs(MovePoint.X).centrepoint.X, Arcs(MovePoint.X).centrepoint.Y, Temp, Dist, Arcs(MovePoint.X).Points(I)
    Next I
    ReDraw
    Moving = False
    Arcs(0).Used = False
  End If
End If

End Sub

Public Sub PutArc(ByVal ColKey As String, Optional JustSel As Boolean = False)
' Allows the dotsshown and centreshown options to go into the PlaceArc function
PlaceArc ColKey, JustSel, -Check1.Value, -Check2.Value
End Sub

Private Sub PlaceArc(ByVal ColKey As String, Optional JustSel As Boolean = False, Optional IncludePoints As Boolean = True, Optional IncludeCentre As Boolean = False)
' Draws an arc on the screen, given its ColKey

Dim nAllThere As Boolean
Dim I As Integer
Dim Arc As Arc
Arc = Arcs(GetIndex(ColKey))
With Arc              ' With the arc

  If .Used = False Then Exit Sub  ' If its not used, get out

  ' Check to make sure that there are 3 points
  For I = 1 To 3
    If .Points(I).Used = False Then nAllThere = True: I = 0: GoTo Justpoint
  Next I

  ' Draw the arc, unless JustSel has been set- i.e., just selections
  If JustSel = False Then I = DrawArcType(Arc, picArc.hDC, 13): picArc.Refresh

Justpoint:

  ' Draw the points, only if wanted, but always if the arc is the selected arc
  For I = 1 To 3
    If .Points(I).Used = True And (IncludePoints = True Or ColKey = CurArcKey) Then
      picArc.ForeColor = IIf(ColKey = CurArcKey, RGB(0, 0, 255), 0)
      picArc.DrawWidth = 3
      picArc.PSet (.Points(I).X, .Points(I).Y)
      picArc.DrawWidth = 1
      picArc.ForeColor = 0
    End If
  Next I

  ' Warning if the points are collinear
  If I = DrawArcResults.COLLINEAR Then
    picArc.CurrentX = 0
    picArc.CurrentY = 0
    MsgBox "Points are collinear; cannot draw arc between points", vbInformation, "Error Creating Arc"
    .Used = False
    Exit Sub
  End If

  If (IncludeCentre = True Or ColKey = CurArcKey) And nAllThere = False Then
    ' Put a dot at the centrepoint, if requested, or if the arc being drawn is
    ' the currently selected arc
    picArc.ForeColor = 255
    picArc.DrawWidth = 3
    picArc.PSet (Arc.centrepoint.X, Arc.centrepoint.Y)
    picArc.DrawWidth = 1
    picArc.ForeColor = 0
  End If

End With

End Sub

' Function to return a new key that can be used for the Arc array
Private Function GetNewKey() As String

Dim I As Integer
Dim J As Integer
Do
  I = I + 1
  For J = 1 To UBound(Arcs)
    If Arcs(J).ColKey = "_" & Trim(Str(I)) Then Exit For
  Next J
  If J = UBound(Arcs) + 1 Then Exit Do
Loop
GetNewKey = "_" & Trim(Str(I))

End Function

' Function to return a new Index that can be used for the Arc array
' If it can't find a spare one, make one!
Private Function GetNewIndex() As Integer

' Remove redundant elements
CompressArray

Dim I As Integer

For I = 1 To UBound(Arcs)
  If Arcs(I).Used = False Then GetNewIndex = I: Exit Function
Next I
ReDim Preserve Arcs(0 To UBound(Arcs) + 1) As Arc
GetNewIndex = UBound(Arcs)
End Function

' Get the index from the key of an arc using the array
Private Function GetIndex(ByVal Key As String) As Integer

Dim I As Integer
Dim J As Integer

For J = 1 To UBound(Arcs)
  If Arcs(J).ColKey = Key Then GetIndex = J: Exit Function
Next J

GetIndex = 0

End Function

' Remove the selection references to an arc
Private Sub Deselect()
If CurArcKey <> "" Then
  CurArcKey = ""
  ReDraw
End If
End Sub

' Place the selection references on an arc
Private Sub SelectArc(ByRef Arc As Arc)
If Arc.Used = False Then Exit Sub
CurArcKey = Arc.ColKey
PutArc CurArcKey, True
End Sub

Private Sub ReDraw()
Dim I As Integer
picArc.Cls
For I = UBound(Arcs) To 1 Step -1
  If Arcs(I).Used = True Then
    PutArc Arcs(I).ColKey
  End If
Next I
End Sub

Private Function FindPoint(ByVal X As Single, ByVal Y As Single) As Point
Dim I As Integer
Dim J As Integer

' Run through all of the points to see if one should be selected
For I = 1 To UBound(Arcs)
  If Arcs(I).Used = True Then
    For J = 1 To 3
      ' If it's used (existant), and within the "Snap" region (Tolerance)
      If Tolerance >= Abs(Arcs(I).Points(J).X - X) And Tolerance >= Abs(Arcs(I).Points(J).Y - Y) Then
        ' Set the point to "found"
        FindPoint.X = I
        FindPoint.Y = J
        ' Get out of here!
        Exit Function
      End If
    Next J
  End If
Next I
' Otherwise, no points found
FindPoint.X = 0
FindPoint.Y = 0

End Function

Private Function FindArc(ByVal X As Single, ByVal Y As Single) As Integer
' Iterate through all of the arcs with 3 steps to see if the click can intercept
' one of them.

Dim I As Integer
Dim Dist As Double
Dim ClickAngle As Double

For I = 1 To UBound(Arcs)
  With Arcs(I)
    ' Step 1. Is it used?
    If .Used = True Then
      ' Step 2. Is the click within the radius of the circle? (+- tolerance?)
      ' Calculate the intersection points
      If CalcIntersectArc(Arcs(I)) = NoError Then
        ' Calculate the radius
        ArcRadius Arcs(I)
        Dist = Sqr((.centrepoint.X - X) ^ 2 + (.centrepoint.Y - Y) ^ 2)
        If Dist < .Radius + Tolerance And Dist > .Radius - Tolerance Then
          ' Step 3. Calculate whether or not the point is within the arc using angles
          ClickAngle = ATNAngle(X - .centrepoint.X, .centrepoint.Y - Y)
          AngleWiseArc Arcs(I)
          If AngleEncompassedArc(Arcs(I), ClickAngle) = True Then
            ' The click is on the arc
            FindArc = I
            Exit Function
          End If
        End If
      End If
    End If
  End With
Next I

End Function

Private Sub CompressArray()
Dim I As Integer
Dim J As Integer
Restart:
For I = 1 To UBound(Arcs)
If Arcs(I).Used = False Then
For J = I + 1 To UBound(Arcs)
Arcs(J - 1) = Arcs(J)
Next J
ReDim Preserve Arcs(0 To UBound(Arcs) - 1) As Arc
GoTo Restart
End If
Next I

End Sub

'
' Returns an angle between 0 and 2*Pi Radians from an X,Y value,
' taking quadrant into account
'
Private Function ATNAngle(ByVal X As Double, ByVal Y As Double) As Double

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

' Bearing and Distance - Return and X,Y point from a point, given bearing, and distance
Public Sub BAndD(ByVal X As Double, ByVal Y As Double, ByVal Angle As Double, _
    ByVal Distance As Double, ByRef OutPoint As Point)

OutPoint.X = X + Distance * Cos(Angle)
OutPoint.Y = Y + Distance * Sin(Angle)

End Sub

' Distance formula
Public Function Distance(ByRef Point1 As Point, ByRef Point2 As Point) As Double
Distance = Sqr((Point2.X - Point1.X) ^ 2 + (Point2.Y - Point1.Y) ^ 2)
End Function
