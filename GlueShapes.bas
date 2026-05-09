' GlueShapes.bas
' VBA macro for Microsoft Visio
'
' Glues two selected shapes along the shortest path.
' Select exactly two shapes before running GlueSelectedShapes.
'
' Algorithm:
'   1. Find the point on shape2 closest to the center of shape1.
'      Add an "Outward" connection point there.
'   2. Find the edge of shape1 closest to that point.
'   3. Build a 1-inch perpendicular from the found point toward shape1's edge.
'   4. Intersect the perpendicular with the nearest edge of shape1.
'   5. Add an "Inward" connection point on shape1 at the intersection.
'   6. Connect the two points with a connector.
'
' Fix for Y-axis bug (step 5/6):
'   Visio stores connection-point coordinates relative to the shape's own
'   local coordinate system (origin = shape pin, i.e. the centre of the
'   bounding box by default).  When the shape is not at the page origin the
'   raw page coordinates must be converted to local shape coordinates before
'   being written to the ConnectionPoints row, otherwise the Y value is wrong.

Option Explicit

' ── public entry point ───────────────────────────────────────────────────────

Public Sub GlueSelectedShapes()
    Dim vsoPage As Visio.Page
    Dim vsoSel  As Visio.Selection
    Dim shp1    As Visio.Shape   ' "first" shape  (connection point goes Inward)
    Dim shp2    As Visio.Shape   ' "second" shape (connection point goes Outward)

    Set vsoPage = ActivePage
    Set vsoSel  = vsoPage.Application.ActiveWindow.Selection

    If vsoSel.Count <> 2 Then
        MsgBox "Please select exactly two shapes.", vbExclamation
        Exit Sub
    End If

    Set shp1 = vsoSel.Item(1)
    Set shp2 = vsoSel.Item(2)

    ' ── Step 1: point on shp2 closest to the centre of shp1 ─────────────────
    Dim cx1 As Double, cy1 As Double
    cx1 = shp1.CellsU("PinX").ResultIU
    cy1 = shp1.CellsU("PinY").ResultIU

    Dim pt2 As PointIU   ' page-space point on shp2
    pt2 = NearestPointOnShape(shp2, cx1, cy1)

    ' Add Outward connection point to shp2
    AddConnectionPoint shp2, pt2.X, pt2.Y, visConPointDirOut

    ' ── Step 2: find the two vertices on shp1 closest to pt2 ────────────────
    Dim v1 As PointIU, v2 As PointIU
    NearestEdgeVertices shp1, pt2.X, pt2.Y, v1, v2

    ' ── Step 3 + 4: perpendicular of length 1" from pt2 toward shp1's edge ──
    Dim edgeAngle  As Double
    Dim perpAngle  As Double
    edgeAngle = Atan2(v2.Y - v1.Y, v2.X - v1.X)

    ' Two candidate perpendicular directions (±90°)
    Dim perpAngleA As Double, perpAngleB As Double
    perpAngleA = edgeAngle + PI / 2
    perpAngleB = edgeAngle - PI / 2

    ' Choose the direction that points toward shp1 centre
    Dim pA As PointIU, pB As PointIU
    pA.X = pt2.X + Cos(perpAngleA)
    pA.Y = pt2.Y + Sin(perpAngleA)
    pB.X = pt2.X + Cos(perpAngleB)
    pB.Y = pt2.Y + Sin(perpAngleB)

    Dim dA As Double, dB As Double
    dA = Distance(pA.X, pA.Y, cx1, cy1)
    dB = Distance(pB.X, pB.Y, cx1, cy1)

    If dA < dB Then
        perpAngle = perpAngleA
    Else
        perpAngle = perpAngleB
    End If

    ' Endpoint of the 1-inch perpendicular (page space)
    Dim perpEnd As PointIU
    perpEnd.X = pt2.X + Cos(perpAngle)
    perpEnd.Y = pt2.Y + Sin(perpAngle)

    ' ── Step 5: intersect perpendicular with the nearest edge of shp1 ────────
    Dim ix As Double, iy As Double
    If Not LineIntersect(pt2.X, pt2.Y, perpEnd.X, perpEnd.Y, _
                         v1.X, v1.Y, v2.X, v2.Y, _
                         ix, iy) Then
        ' Fallback: project pt2 perpendicularly onto the edge segment
        ProjectPointOnSegment pt2.X, pt2.Y, v1.X, v1.Y, v2.X, v2.Y, ix, iy
    End If

    ' ── Step 6: add Inward connection point to shp1 (page → local coords) ───
    '   Convert page-space intersection to shp1 local coordinates.
    '   Visio's XYToLocal / LocalToXY helpers work in the shape's local frame.
    Dim localX As Double, localY As Double
    PageToLocalCoords shp1, ix, iy, localX, localY

    AddConnectionPointLocal shp1, localX, localY, visConPointDirIn

    ' ── Step 7: drop a connector and glue the two points ────────────────────
    ConnectPoints vsoPage, shp1, shp2

    MsgBox "Shapes glued successfully.", vbInformation
End Sub

' ── helpers ──────────────────────────────────────────────────────────────────

Private Type PointIU
    X As Double
    Y As Double
End Type

Private Const PI As Double = 3.14159265358979

' Returns the point on the bounding-box perimeter of shp that is closest
' to (targetX, targetY) in page-space internal units.
Private Function NearestPointOnShape(shp As Visio.Shape, _
                                     targetX As Double, targetY As Double) As PointIU
    Dim bLeft   As Double, bRight  As Double
    Dim bBottom As Double, bTop    As Double

    bLeft   = shp.CellsU("PinX").ResultIU - shp.CellsU("Width").ResultIU  / 2
    bRight  = shp.CellsU("PinX").ResultIU + shp.CellsU("Width").ResultIU  / 2
    bBottom = shp.CellsU("PinY").ResultIU - shp.CellsU("Height").ResultIU / 2
    bTop    = shp.CellsU("PinY").ResultIU + shp.CellsU("Height").ResultIU / 2

    ' Clamp target to bounding box → closest surface point
    Dim cx As Double, cy As Double
    cx = ClampD(targetX, bLeft, bRight)
    cy = ClampD(targetY, bBottom, bTop)

    ' If the target is inside the shape, project it onto the nearest edge
    If targetX >= bLeft And targetX <= bRight And _
       targetY >= bBottom And targetY <= bTop Then
        Dim dL As Double, dR As Double, dB As Double, dT As Double
        dL = targetX - bLeft
        dR = bRight  - targetX
        dB = targetY - bBottom
        dT = bTop    - targetY
        Dim dMin As Double
        dMin = MinD(MinD(dL, dR), MinD(dB, dT))
        Select Case dMin
            Case dL: cx = bLeft:   cy = targetY
            Case dR: cx = bRight:  cy = targetY
            Case dB: cy = bBottom: cx = targetX
            Case dT: cy = bTop:    cx = targetX
        End Select
    End If

    NearestPointOnShape.X = cx
    NearestPointOnShape.Y = cy
End Function

' Finds the two vertices of the bounding box of shp whose edge is closest
' to (nearX, nearY).  Returns them in v1/v2.
Private Sub NearestEdgeVertices(shp As Visio.Shape, _
                                nearX As Double, nearY As Double, _
                                ByRef v1 As PointIU, ByRef v2 As PointIU)
    Dim bLeft   As Double, bRight  As Double
    Dim bBottom As Double, bTop    As Double

    bLeft   = shp.CellsU("PinX").ResultIU - shp.CellsU("Width").ResultIU  / 2
    bRight  = shp.CellsU("PinX").ResultIU + shp.CellsU("Width").ResultIU  / 2
    bBottom = shp.CellsU("PinY").ResultIU - shp.CellsU("Height").ResultIU / 2
    bTop    = shp.CellsU("PinY").ResultIU + shp.CellsU("Height").ResultIU / 2

    ' Measure distance from nearX/nearY to each of the four edges
    Dim dL As Double, dR As Double, dB As Double, dT As Double
    dL = DistPointToSegment(nearX, nearY, bLeft, bBottom, bLeft, bTop)
    dR = DistPointToSegment(nearX, nearY, bRight, bBottom, bRight, bTop)
    dB = DistPointToSegment(nearX, nearY, bLeft, bBottom, bRight, bBottom)
    dT = DistPointToSegment(nearX, nearY, bLeft, bTop, bRight, bTop)

    Dim dMin As Double
    dMin = MinD(MinD(dL, dR), MinD(dB, dT))

    Select Case dMin
        Case dL
            v1.X = bLeft:  v1.Y = bBottom
            v2.X = bLeft:  v2.Y = bTop
        Case dR
            v1.X = bRight: v1.Y = bBottom
            v2.X = bRight: v2.Y = bTop
        Case dB
            v1.X = bLeft:  v1.Y = bBottom
            v2.X = bRight: v2.Y = bBottom
        Case dT
            v1.X = bLeft:  v1.Y = bTop
            v2.X = bRight: v2.Y = bTop
    End Select
End Sub

' Converts page-space coordinates to the shape's local coordinate system.
' In Visio, connection-point X/Y cells are in the shape's local frame whose
' origin is the shape's pin (PinX, PinY) and whose axes are rotated by the
' shape's angle.
Private Sub PageToLocalCoords(shp As Visio.Shape, _
                               pageX As Double, pageY As Double, _
                               ByRef localX As Double, ByRef localY As Double)
    Dim pinX  As Double, pinY  As Double, ang As Double
    pinX = shp.CellsU("PinX").ResultIU
    pinY = shp.CellsU("PinY").ResultIU
    ang  = shp.CellsU("Angle").ResultIU   ' radians

    ' Translate to pin origin
    Dim dx As Double, dy As Double
    dx = pageX - pinX
    dy = pageY - pinY

    ' Rotate by -angle to get local axes
    localX = dx * Cos(-ang) - dy * Sin(-ang)
    localY = dx * Sin(-ang) + dy * Cos(-ang)
End Sub

' Adds a connection point to shp using PAGE-space coordinates.
' Internally converts to local coords before writing to ShapeSheet.
Private Sub AddConnectionPoint(shp As Visio.Shape, _
                                pageX As Double, pageY As Double, _
                                direction As Integer)
    Dim localX As Double, localY As Double
    PageToLocalCoords shp, pageX, pageY, localX, localY
    AddConnectionPointLocal shp, localX, localY, direction
End Sub

' Adds a connection point using LOCAL (shape-space) coordinates.
Private Sub AddConnectionPointLocal(shp As Visio.Shape, _
                                     localX As Double, localY As Double, _
                                     direction As Integer)
    Dim cpRow As Long
    cpRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagCnnctPt)

    shp.CellsSRC(visSectionConnectionPts, cpRow, visCnnctX).ResultIU = localX
    shp.CellsSRC(visSectionConnectionPts, cpRow, visCnnctY).ResultIU = localY
    shp.CellsSRC(visSectionConnectionPts, cpRow, visCnnctDirX).FormulaU = "0"
    shp.CellsSRC(visSectionConnectionPts, cpRow, visCnnctDirY).FormulaU = "0"
    shp.CellsSRC(visSectionConnectionPts, cpRow, visCnnctType).FormulaU = CStr(direction)
End Sub

' Drops a 1D connector on the page and glues its endpoints to the last
' connection points of shp1 and shp2.
Private Sub ConnectPoints(pg As Visio.Page, _
                           shp1 As Visio.Shape, shp2 As Visio.Shape)
    Dim connector As Visio.Shape
    Set connector = pg.Drop(pg.Application.ConnectorToolDataObject, 0, 0)

    ' Glue begin-point to last connection point of shp1 (Inward)
    Dim cpCount1 As Long
    cpCount1 = shp1.RowCount(visSectionConnectionPts)
    connector.CellsU("BeginX").GlueTo _
        shp1.CellsSRC(visSectionConnectionPts, cpCount1 - 1, visCnnctX)

    ' Glue end-point to last connection point of shp2 (Outward)
    Dim cpCount2 As Long
    cpCount2 = shp2.RowCount(visSectionConnectionPts)
    connector.CellsU("EndX").GlueTo _
        shp2.CellsSRC(visSectionConnectionPts, cpCount2 - 1, visCnnctX)
End Sub

' ── geometry helpers ─────────────────────────────────────────────────────────

Private Function Distance(x1 As Double, y1 As Double, _
                           x2 As Double, y2 As Double) As Double
    Distance = Sqr((x2 - x1) ^ 2 + (y2 - y1) ^ 2)
End Function

Private Function DistPointToSegment(px As Double, py As Double, _
                                     ax As Double, ay As Double, _
                                     bx As Double, by As Double) As Double
    Dim projX As Double, projY As Double
    ProjectPointOnSegment px, py, ax, ay, bx, by, projX, projY
    DistPointToSegment = Distance(px, py, projX, projY)
End Function

' Projects point (px, py) onto segment (ax,ay)-(bx,by); result in (rx, ry).
Private Sub ProjectPointOnSegment(px As Double, py As Double, _
                                   ax As Double, ay As Double, _
                                   bx As Double, by As Double, _
                                   ByRef rx As Double, ByRef ry As Double)
    Dim dx As Double, dy As Double
    dx = bx - ax
    dy = by - ay
    Dim lenSq As Double
    lenSq = dx * dx + dy * dy
    If lenSq = 0 Then
        rx = ax: ry = ay
        Exit Sub
    End If
    Dim t As Double
    t = ((px - ax) * dx + (py - ay) * dy) / lenSq
    t = ClampD(t, 0, 1)
    rx = ax + t * dx
    ry = ay + t * dy
End Sub

' Computes intersection of infinite lines through (x1,y1)-(x2,y2) and
' (x3,y3)-(x4,y4).  Returns True and sets (ix,iy) if lines are not parallel.
Private Function LineIntersect(x1 As Double, y1 As Double, _
                                x2 As Double, y2 As Double, _
                                x3 As Double, y3 As Double, _
                                x4 As Double, y4 As Double, _
                                ByRef ix As Double, ByRef iy As Double) As Boolean
    Dim denom As Double
    denom = (x1 - x2) * (y3 - y4) - (y1 - y2) * (x3 - x4)
    If Abs(denom) < 0.000000000001 Then
        LineIntersect = False
        Exit Function
    End If
    Dim t As Double
    t = ((x1 - x3) * (y3 - y4) - (y1 - y3) * (x3 - x4)) / denom
    ix = x1 + t * (x2 - x1)
    iy = y1 + t * (y2 - y1)
    LineIntersect = True
End Function

Private Function Atan2(y As Double, x As Double) As Double
    If x > 0 Then
        Atan2 = Atn(y / x)
    ElseIf x < 0 And y >= 0 Then
        Atan2 = Atn(y / x) + PI
    ElseIf x < 0 And y < 0 Then
        Atan2 = Atn(y / x) - PI
    ElseIf x = 0 And y > 0 Then
        Atan2 = PI / 2
    ElseIf x = 0 And y < 0 Then
        Atan2 = -PI / 2
    Else
        Atan2 = 0
    End If
End Function

Private Function ClampD(v As Double, lo As Double, hi As Double) As Double
    If v < lo Then
        ClampD = lo
    ElseIf v > hi Then
        ClampD = hi
    Else
        ClampD = v
    End If
End Function

Private Function MinD(a As Double, b As Double) As Double
    If a < b Then MinD = a Else MinD = b
End Function
