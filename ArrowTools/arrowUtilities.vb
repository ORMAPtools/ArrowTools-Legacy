Imports System.Runtime.InteropServices
Imports System.Drawing
Imports ESRI.ArcGIS.Geometry
Imports System.Xml
Imports System.IO
Imports ESRI.ArcGIS.Editor
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Display
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.esriSystem
Imports System.Windows.Forms
Imports System.Xml.Linq
Imports System.Linq
Imports ESRI.ArcGIS.Framework

Module arrowUtilities

#Region "Module Variables"
    ''' <summary>
    ''' Enumeration of arrow types.
    ''' </summary>
    Public Enum arrowCategories
        Straight = 0
        LandHook = 1
        NoDashes = 2
        OneDash = 3
        TwoDashes = 4
        ThreeDashes = 5
        FourDashes = 6
        SingleArrow = 7
        RoadTic = 8
    End Enum

    ''' <summary>
    ''' Enumeration of single arrow styles
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum arrowStyles
        Straight = 0
        Leader = 1
        Zigzag = 2
        Freeform = 3
        RoadTic = 4
    End Enum

    Structure arrowType
        Dim category As Integer
        Dim style As Integer
    End Structure

#Region "constants"
    Friend Const SEPARATOR = "-"
    Friend Const SHORTER = "Shorter - Space"
    Friend Const LONGER = "Longer - Shift Space"
    Friend Const FLIP = "Flip Arrows - F"
    Friend Const UNLOCK = "Unlock/Lock Angle - U"
    Friend Const SWITCH = "Switch Arrowheads - S"
    Friend Const CANCEL = "Cancel - Esc"
    Friend Const SCALE10 = "10 Scale"
    Friend Const SCALE20 = "20 Scale"
    Friend Const SCALE30 = "30 Scale"
    Friend Const SCALE40 = "40 Scale"
    Friend Const SCALE50 = "50 Scale"
    Friend Const SCALE100 = "100 Scale"
    Friend Const SCALE200 = "200 Scale"
    Friend Const SCALE400 = "400 Scale"
    Friend Const SCALE800 = "800 Scale"
    Friend Const SCALE1000 = "1000 Scale"
    Friend Const SCALE2000 = "2000 Scale"
    Friend Const NARROWER = "Narrower"
    Friend Const WIDER = "Wider"
    Friend Const TOPOINT = "Slide toward point"
    Friend Const TOEND = "Slide toward end"
    Friend Const CURVELESS = "Less curve"
    Friend Const CURVEMORE = "More curve"
    Friend Const STYLE_STRAIGHT = "Straight arrow style"
    Friend Const STYLE_LEADER = "Leader arrow style"
    Friend Const STYLE_ZIGZAG = "Zigzag arrow style"
    Friend Const STYLE_FREEFORM = "Freeform arrow style"
    Friend Const STYLE_ROADTIC = "Road Tic"
    Friend Const SAVEDEFAULT = "Save as default zigzag arrow"
    Friend Const HELP = "Help"
    Friend Const FINISH = "Finish Arrow"
    Friend Const SELECTION = "Select new arrows"

#End Region

    Friend WithEvents _editEvents As Editor
    Friend _editor As IEditor3
    Friend _firstPoint As IPoint = New ESRI.ArcGIS.Geometry.Point
    Friend _flipArrows As Boolean = False
    Friend _arrowScale As Double
    Friend _arrowAngle As Double
    Friend _angleIsSet As Boolean
    Friend _featureClass As IFeatureLayer
    Friend _featureTemplate As IEditTemplate
    Friend _lastPoint As IPoint = New ESRI.ArcGIS.Geometry.Point
    Friend _prevPoint As IPoint = New ESRI.ArcGIS.Geometry.Point
    Friend _pointNumber As Integer = 1
    Friend _menuitems As List(Of String) = New List(Of String)
    Friend _arrowheadIsSwitched As Boolean = False
    Friend _arrowOffset As Double
    Friend _installationFolder As String
    Friend _zigzagWidth As Double = CDbl(GetSetting("OR_DOR_dimensionArrows", "default", "zigzagWidth", "5"))
    Friend _zigzagCurve As Double = CDbl(GetSetting("OR_DOR_dimensionArrows", "default", "zigzagCurve", "5"))
    Friend _zigzagPosition As Double = CDbl(GetSetting("OR_DOR_dimensionArrows", "default", "zigzagPosition", "10"))
    Friend _thisArrow As arrowType
    Friend _freeformPoints() As IPoint
    Friend _graphicElement(1) As IElement
    Friend _mxDoc As IMxDocument = My.ArcMap.Document
    Friend _OIDlist As New List(Of Integer)
    Friend _selectNewArrows As Boolean
    Friend _snapPoint As IPoint = New ESRI.ArcGIS.Geometry.Point

#End Region

    ''' <summary>
    ''' Set up the arrow display after the first mouse click
    ''' </summary>
    ''' <param name="thePoint">Mouse position in data frame coordinates</param>
    ''' <remarks></remarks>
    Friend Sub setLineFeedback(ByVal thePoint As IPoint)
        Try
            _firstPoint.PutCoords(thePoint.X, thePoint.Y)
            'Debug.Print("setFeedback - " & CStr(_snapPoint.X) & ", " & CStr(thePoint.X))
            If _thisArrow.style = arrowStyles.Freeform Then
                ReDim _freeformPoints(0)
                _freeformPoints(0) = New ESRI.ArcGIS.Geometry.Point
                _freeformPoints(0).PutCoords(thePoint.X, thePoint.Y)
            End If

            If Not _thisArrow.category = arrowCategories.SingleArrow Then
                SetScale(_firstPoint)
                _arrowAngle = perpendicularAngle()
                _angleIsSet = _arrowAngle <> Nothing
               End If
        Catch ex As Exception
            MessageBox.Show("setLineFeedback - " & ex.ToString)
        End Try
    End Sub

    ''' <summary>
    ''' Places the arrows after the last mouse click
    ''' </summary>
    ''' <param name="endPoint">Mouse position in data frame coordinates</param>
    ''' <param name="middlePoint">Optional second mouse point for arrows that 
    ''' require three points</param>
    ''' <remarks></remarks>
    Friend Sub placeArrows(ByVal endPoint As IPoint, Optional ByVal middlePoint As IPoint = Nothing)
        Try
            Dim editorUID As New UID
            Dim theEditSketch As IEditSketch2

            editorUID.Value = "esriEditor.Editor"
            theEditSketch = DirectCast(My.ArcMap.Application.FindExtensionByCLSID(editorUID), IEditSketch2)

            theEditSketch.GeometryType = esriGeometryType.esriGeometryPolyline

            Dim count As Integer
            Dim theEditTask As IEditTask
            Dim taskName As String
            taskName = _editor.CurrentTask.Name
            For count = 0 To _editor.TaskCount - 1
                theEditTask = _editor.Task(count)
                If theEditTask.Name = "Create New Feature" Then
                    _editor.CurrentTask = theEditTask
                    Exit For
                End If
            Next count

            If _thisArrow.category = arrowCategories.SingleArrow Then
                Dim polyline As IPolyline = getSingleArrowGeometry(endPoint, middlePoint)
                theEditSketch.Geometry = polyline
                theEditSketch.FinishSketch()
                _OIDlist.Add(_editor.EditSelection.Next.OID)
            Else
                Dim polyLineArray As IPolylineArray = getArrowGeometry(endPoint)
                For count = 0 To polyLineArray.Count - 1
                    theEditSketch.Geometry = polyLineArray.Element(count)
                    theEditSketch.FinishSketch()
                    _OIDlist.Add(_editor.EditSelection.Next.OID)
                Next
            End If

            For count = 0 To _editor.TaskCount - 1
                theEditTask = _editor.Task(count)
                If theEditTask.Name = taskName Then
                    _editor.CurrentTask = theEditTask
                    Exit For
                End If
            Next count

            clearAll()
            _snapPoint = Nothing
            selectNewArrows()
            Catch ex As Exception
            MessageBox.Show("placeArrows - " & ex.ToString)
        End Try
    End Sub

    Friend Sub checkSnap()
        If Not _snapPoint Is Nothing Then
            _editor.InvertAgent(_snapPoint, 0)
        End If

        Dim snap As ISnapEnvironment = DirectCast(_editor, ISnapEnvironment)
        
        Dim isSnap As Boolean
        _snapPoint = _mxDoc.CurrentLocation
        Debug.Print("before - " & CStr(_snapPoint.X))
        isSnap = snap.SnapPoint(_snapPoint)
        Debug.Print("after - " & CStr(_snapPoint.X))
        If isSnap Then
            _editor.InvertAgent(_snapPoint, 0)
            Debug.Print("second - " & CStr(_snapPoint.X))
        End If
    End Sub
    ''' <summary>
    ''' Adds newly created items to the selection set so that common attributes can be edited easily
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub selectNewArrows()
        If _selectNewArrows Then
            Dim curLayer As ILayer = _editor.CurrentTemplate.Layer
            Dim featureSelection As IFeatureSelection = DirectCast(curLayer, IFeatureSelection)
            featureSelection.CombinationMethod = esriSelectionResultEnum.esriSelectionResultAdd
            For Each OID As Integer In _OIDlist
                featureSelection.SelectionSet.Add(OID)
            Next
            Dim selectionEvents As ISelectionEvents = DirectCast(My.Document.FocusMap, ISelectionEvents)
            selectionEvents.SelectionChanged()
        Else
            _mxDoc.FocusMap.ClearSelection()
        End If
    End Sub

    ''' <summary>
    ''' Clears the list of objectIDs used for selection when the selection set is cleared
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub EditorEvents_onSelectionChanged() Handles _editEvents.OnSelectionChanged
        If _editor.SelectionCount = 0 Then
            _OIDlist.Clear()
        End If
    End Sub


    ''' <summary>
    ''' Finds the angle perpendicular to the selected line for straight opposing arrows
    ''' </summary>
    ''' <returns>the perpendicular angle</returns>
    ''' <remarks></remarks>
    Friend Function perpendicularAngle() As Double
        Try
            Dim theGeometry As ESRI.ArcGIS.Geometry.IGeometry
            Dim theEnvelope As ESRI.ArcGIS.Geometry.IEnvelope
            Dim theSpatialFilter As ESRI.ArcGIS.Geodatabase.ISpatialFilter
            Dim theFeatureClass As ESRI.ArcGIS.Geodatabase.IFeatureClass
            Dim theEditLayers As ESRI.ArcGIS.Editor.IEditLayers
            Dim theFeatureCursor As ESRI.ArcGIS.Geodatabase.IFeatureCursor
            Dim theFeature As ESRI.ArcGIS.Geodatabase.IFeature
            Dim theMap As ESRI.ArcGIS.Carto.IMap
            Dim ShapeFieldName As String

            perpendicularAngle = Nothing

            'Get the Map from the editor
            theEditLayers = TryCast(_editor, ESRI.ArcGIS.Editor.IEditLayers)
            theMap = TryCast(_editor.Map, ESRI.ArcGIS.Carto.IMap)

            'Pass point to CreateSearchShape which creates a geometry around the point
            'The larger geometry is an envelope and will give us better search results
            'The click therefore doesn't have to be exactly on the feature
            theGeometry = _editor.CreateSearchShape(_firstPoint)
            theEnvelope = TryCast(theGeometry, ESRI.ArcGIS.Geometry.IEnvelope)

            'Create a new spatial filter and use the new envelope as the geometry
            theSpatialFilter = New ESRI.ArcGIS.Geodatabase.SpatialFilter
            theSpatialFilter.Geometry = theEnvelope
            ShapeFieldName = theEditLayers.CurrentLayer.FeatureClass.ShapeFieldName
            theSpatialFilter.OutputSpatialReference(ShapeFieldName) = theMap.SpatialReference
            theSpatialFilter.GeometryField = _
             theEditLayers.CurrentLayer.FeatureClass.ShapeFieldName
            theSpatialFilter.SpatialRel = esriSpatialRelEnum.esriSpatialRelIntersects

            Dim enumLayers As IEnumLayer
            Dim eachLayer As IFeatureLayer
            Dim pUID As UID = New UIDClass()

            pUID.Value = "{40A9E885-5533-11d0-98BE-00805F7CED21}" 'only select IfeatureLayer

            'Get all the feature layers from the map
            enumLayers = theMap.Layers(pUID, True)
            enumLayers.Reset()
            eachLayer = DirectCast(enumLayers.Next, IFeatureLayer)

            Dim featureSet As IFeature = New Feature
            Dim editLayer As IEditLayers = DirectCast(_editor, IEditLayers)

            Do Until eachLayer Is Nothing
                'Only search for lines or polygons
                If eachLayer.Valid Then
                    If eachLayer.FeatureClass.ShapeType = _
                     esriGeometryType.esriGeometryPolyline Then
                        theFeatureClass = eachLayer.FeatureClass
                        theFeatureCursor = theFeatureClass.Search(theSpatialFilter, False) 'Do the search
                        theFeature = theFeatureCursor.NextFeature 'Get the first feature
                        If Not theFeature Is Nothing Then
                            Dim polyLine As IPolyline
                            polyLine = DirectCast(theFeature.Shape, IPolyline)
                            Dim fromPoint As ESRI.ArcGIS.Geometry.Point = New ESRI.ArcGIS.Geometry.Point
                            Dim toPoint As ESRI.ArcGIS.Geometry.Point = New ESRI.ArcGIS.Geometry.Point
                            Dim pointDist As Double
                            Dim perpendicularLine As ILine = New Line

                            polyLine.QueryPointAndDistance( _
                                esriSegmentExtension.esriExtendAtFrom, _firstPoint, False, _
                                fromPoint, pointDist, Nothing, False)
                            polyLine.QueryNormal(esriSegmentExtension.esriExtendAtFrom, _
                                pointDist, False, 50, perpendicularLine)
                            perpendicularLine.QueryFromPoint(fromPoint)
                            perpendicularLine.QueryToPoint(toPoint)
                            perpendicularAngle = perpendicularLine.Angle
                            Exit Function
                        End If
                    End If
                End If
                eachLayer = DirectCast(enumLayers.Next, IFeatureLayer)
            Loop

        Catch ex As Exception
            MsgBox("perpendicularAngle" & vbCrLf & ex.Message, MsgBoxStyle.OkOnly, "Error")
        End Try
    End Function

    ''' <summary>
    ''' Shows the help file
    ''' </summary>
    ''' <remarks></remarks>
    Friend Sub showHelp()
        Dim myproc As System.Diagnostics.Process = New System.Diagnostics.Process
        myproc.EnableRaisingEvents = False
        myproc.StartInfo.FileName = _installationFolder & "\ArrowTools.chm"
        myproc.Start()
    End Sub

    ''' <summary>
    ''' Finds the map scale based on the underlying MapIndex polygon and saves it in
    ''' the _arrowScale variable
    ''' </summary>
    ''' <param name="thePoint">The first arrow point</param>
    ''' <remarks>If there is no map index a scale of 1 (1"=100') is set</remarks>
    Friend Sub SetScale(ByVal thePoint As IPoint)
        Try
            Dim theGeometry As ESRI.ArcGIS.Geometry.IGeometry
            Dim theEnvelope As ESRI.ArcGIS.Geometry.IEnvelope
            Dim theSpatialFilter As ESRI.ArcGIS.Geodatabase.ISpatialFilter
            Dim theFeatureClass As ESRI.ArcGIS.Geodatabase.IFeatureClass
            Dim theEditLayers As ESRI.ArcGIS.Editor.IEditLayers
            Dim theFeatureCursor As ESRI.ArcGIS.Geodatabase.IFeatureCursor
            Dim theFeature As ESRI.ArcGIS.Geodatabase.IFeature
            Dim theMap As ESRI.ArcGIS.Carto.IMap
            Dim ShapeFieldName As String

            'Get the Map from the editor
            theEditLayers = TryCast(_editor, ESRI.ArcGIS.Editor.IEditLayers)
            theMap = TryCast(_editor.Map, ESRI.ArcGIS.Carto.IMap)

            'Pass point to CreateSearchShape which creates a geometry around the point
            'The larger geometry is an envelope and will give us better search results
            'The click therefore doesn't have to be exactly on the feature
            theGeometry = _editor.CreateSearchShape(thePoint)
            theEnvelope = TryCast(theGeometry, ESRI.ArcGIS.Geometry.IEnvelope)

            'Create a new spatial filter and use the new envelope as the geometry
            theSpatialFilter = New ESRI.ArcGIS.Geodatabase.SpatialFilter
            theSpatialFilter.Geometry = theEnvelope
            ShapeFieldName = theEditLayers.CurrentLayer.FeatureClass.ShapeFieldName
            theSpatialFilter.OutputSpatialReference(ShapeFieldName) = theMap.SpatialReference
            theSpatialFilter.GeometryField = theEditLayers.CurrentLayer.FeatureClass.ShapeFieldName
            theSpatialFilter.SpatialRel = esriSpatialRelEnum.esriSpatialRelIntersects

            Dim enumLayers As IEnumLayer
            Dim eachLayer As IFeatureLayer
            Dim pUID As UID = New UIDClass()

            pUID.Value = "{40A9E885-5533-11d0-98BE-00805F7CED21}" 'only select IfeatureLayer

            'Get all the feature layers from the map
            enumLayers = theMap.Layers(pUID, True)
            enumLayers.Reset()
            eachLayer = DirectCast(enumLayers.Next, IFeatureLayer)

            Do Until eachLayer Is Nothing
                If InStr(UCase(eachLayer.Name), "MAPINDEX") > 0 Then
                    'Only search the specified geometry type
                    If eachLayer.FeatureClass.ShapeType = _
                        esriGeometryType.esriGeometryPolygon Then
                        theFeatureClass = eachLayer.FeatureClass
                        theFeatureCursor = theFeatureClass.Search(theSpatialFilter, False) 'Do the search
                        theFeature = theFeatureCursor.NextFeature 'Get the first feature
                        If Not theFeature Is Nothing Then
                            _arrowScale = _
                                CDbl(theFeature.Value(theFeature.Fields.FindField("MapScale"))) / 1200
                            Exit Sub
                        End If
                    End If
                End If
                eachLayer = DirectCast(enumLayers.Next, IFeatureLayer)
            Loop
            'if there is no map index then set it to 100 scale
            _arrowScale = 1
        Catch ex As Exception
            MsgBox("SetScale" & vbCrLf & ex.Message, MsgBoxStyle.OkOnly, "Error")
        End Try
    End Sub

    ''' <summary>
    ''' Get the geometry of the paired arrows or road tics
    ''' </summary>
    ''' <param name="endPoint">Mouse position in data frame coordinates</param>
    ''' <returns>An IPolyLineArray of two elements</returns>
    ''' <remarks></remarks>
    Friend Function getArrowGeometry(ByVal endPoint As IPoint) As IPolylineArray
        getArrowGeometry = Nothing
        If checkXML() Then
            Try

                If _angleIsSet Then
                    endPoint.ConstrainAngle(_arrowAngle, _firstPoint, True)
                End If

                Dim count As Integer

                Dim geoString As String
                geoString = ReadXML(_thisArrow.category)
                Dim segmentCount As Integer
                segmentCount = CInt(Left(geoString, InStr(geoString, ",") - 1))
                Dim segmentString() As String = geoString.Split(CChar(","))

                Dim segment1() As ILine
                Dim segment2() As ILine
                Dim missing As Object = Type.Missing

                Dim arrow1 As IPolyline = New Polyline
                Dim arrow2 As IPolyline = New Polyline

                'Draw the left arrow
                Dim path1 As ISegmentCollection = New Polyline
                Dim path2 As ISegmentCollection = New Polyline

                Dim scaleFactor As Double
                If _pointNumber = 3 Then
                    scaleFactor = _arrowOffset / CDbl(segmentString(UBound(segmentString))) / _arrowScale
                Else
                    scaleFactor = 1
                End If

                ReDim segment1(segmentCount - 1)
                For count = 0 To segmentCount - 1
                    segment1(count) = New ESRI.ArcGIS.Geometry.Line
                Next

                Dim coordStep As Integer = 1
                For count = 0 To segmentCount - 1
                    Dim point1 As IPoint = New ESRI.ArcGIS.Geometry.Point
                    Dim point2 As IPoint = New ESRI.ArcGIS.Geometry.Point
                    point1.X = CDbl(segmentString(coordStep))
                    If count = 0 Then
                        point1.Y = CDbl(segmentString(coordStep + 1))
                    Else
                        point1.Y = CDbl(segmentString(coordStep + 1)) * scaleFactor
                    End If
                    segment1(count).FromPoint = point1
                    point2.X = CDbl(segmentString(coordStep + 2))
                    point2.Y = CDbl(segmentString(coordStep + 3)) * scaleFactor
                    segment1(count).ToPoint = point2
                    path1.AddSegment(DirectCast(segment1(count), ISegment), missing, missing)
                    coordStep = coordStep + 4
                Next

                'Draw the right arrow
                ReDim segment2(segmentCount - 1)
                For count = 0 To segmentCount - 1
                    segment2(count) = New ESRI.ArcGIS.Geometry.Line
                Next

                coordStep = 1
                For count = 0 To segmentCount - 1
                    Dim point1 As IPoint = New ESRI.ArcGIS.Geometry.Point
                    Dim point2 As IPoint = New ESRI.ArcGIS.Geometry.Point
                    point1.X = CDbl(segmentString(coordStep)) * -1
                    If count = 0 Then
                        point1.Y = CDbl(segmentString(coordStep + 1))
                    Else
                        point1.Y = CDbl(segmentString(coordStep + 1)) * scaleFactor
                    End If
                    segment2(count).FromPoint = point1
                    point2.X = CDbl(segmentString(coordStep + 2)) * -1
                    point2.Y = CDbl(segmentString(coordStep + 3)) * scaleFactor
                    If _thisArrow.category = arrowCategories.LandHook And count = segmentCount - 1 Then
                        point2.Y = point2.Y * -1
                    End If
                    segment2(count).ToPoint = point2
                    path2.AddSegment(DirectCast(segment2(count), ISegment), missing, missing)
                    coordStep = coordStep + 4
                Next

                arrow1 = DirectCast(path1, IPolyline)
                arrow2 = DirectCast(path2, IPolyline)

                If _arrowheadIsSwitched Then
                    arrow1.ReverseOrientation()
                    arrow2.ReverseOrientation()
                End If

                If _thisArrow.category = arrowCategories.Straight Then
                    If _flipArrows Then
                        arrow1 = DirectCast(path2, IPolyline)
                        arrow2 = DirectCast(path1, IPolyline)
                    End If
                End If

                'use a new line to get the angle of the vector between the points
                Dim vector As ILine = New Line
                Dim vectorFrom As IPoint = New ESRI.ArcGIS.Geometry.Point
                Dim vectorTo As IPoint = New ESRI.ArcGIS.Geometry.Point

                vectorFrom.PutCoords(_firstPoint.X, _firstPoint.Y)
                vectorTo.PutCoords(endPoint.X, endPoint.Y)
                vector.FromPoint = vectorFrom
                vector.ToPoint = vectorTo
                Dim vectorAngle As Double = vector.Angle
                vector = Nothing

                'Transform the polylines by moving, rotating and scaling them to the proper positions
                Dim transform As ITransform2D = DirectCast(arrow1, ITransform2D)
                transform.Move(_firstPoint.X, _firstPoint.Y)
                transform.Rotate(_firstPoint, vectorAngle)
                transform.Scale(_firstPoint, _arrowScale, _arrowScale)

                transform = DirectCast(arrow2, ITransform2D)
                transform.Move(endPoint.X, endPoint.Y)
                transform.Rotate(endPoint, vectorAngle)
                transform.Scale(endPoint, _arrowScale, _arrowScale)

                Dim arrowArray As IPolylineArray = New PolylineArray
                arrowArray.Add(arrow1)
                If Not _thisArrow.category = arrowCategories.RoadTic Then
                    arrowArray.Add(arrow2)
                End If

                Return arrowArray
            Catch ex As Exception
                MessageBox.Show("getArrowGeometry - " & ex.ToString)
            End Try
        End If
    End Function

    ''' <summary>
    ''' Show the arrows while moving the mouse
    ''' </summary>
    ''' <param name="endPoint">Mouse position in data frame coordinates</param>
    ''' <param name="middlePoint">Option second mouse point for arrows that 
    ''' require three points</param>
    ''' <remarks></remarks>
    Friend Sub showLineFeedback(ByVal endPoint As IPoint, Optional ByVal middlePoint As IPoint = Nothing)
        If checkXML() Then
            Try
                _lastPoint.PutCoords(endPoint.X, endPoint.Y)

                Dim theGraphicsContainer As IGraphicsContainer = My.ArcMap.Document.ActiveView.GraphicsContainer

                Try
                    theGraphicsContainer.DeleteElement(_graphicElement(0))
                    theGraphicsContainer.DeleteElement(_graphicElement(1))
                Catch
                    'Debug.Print("No graphic element to delete")
                End Try

                If _thisArrow.category = arrowCategories.SingleArrow Then
                    Dim polyLine As IPolyline = getSingleArrowGeometry(endPoint, middlePoint)
                    drawArrowImage(polyLine, 0)
                Else
                    Dim polyLineArray As IPolylineArray = getArrowGeometry(endPoint)
                    drawArrowImage(polyLineArray.Element(0), 0)
                    If Not _thisArrow.category = arrowCategories.RoadTic Then
                        drawArrowImage(polyLineArray.Element(1), 1)
                    End If
                End If
            Catch ex As Exception
                MessageBox.Show("showLineFeedback - " & ex.ToString)
            End Try
        End If
    End Sub

    ''' <summary>
    ''' Get the geometry for a single arrow
    ''' </summary>
    ''' <param name="endPoint">Mouse position in data frame coordinates</param>
    ''' <param name="middlePoint">Optional second mouse point for arrows that 
    ''' require three points</param>
    ''' <returns>The arrow geometry as an IPolyLine</returns>
    ''' <remarks></remarks>
    Friend Function getSingleArrowGeometry(ByVal endPoint As IPoint, ByVal middlePoint As IPoint) As IPolyline
        Dim thePolyline As IPolyline = New Polyline
        If checkXML() Then
            Try
                Dim count As Integer
                Dim segpoints(3) As IPoint

                If Not _thisArrow.style = arrowStyles.Freeform Then
                    If _angleIsSet And _pointNumber = 3 Then
                        endPoint.ConstrainAngle(0, middlePoint, True)
                    End If

                    If middlePoint Is Nothing Then
                        middlePoint = New ESRI.ArcGIS.Geometry.Point
                        middlePoint.PutCoords(endPoint.X, endPoint.Y)
                    End If

                    For count = 0 To 3
                        segpoints(count) = New ESRI.ArcGIS.Geometry.Point
                    Next
                End If

                Dim offset As Integer = 1
                If _flipArrows Then
                    offset = -1
                End If

                Dim missing As Object = Type.Missing

                Select Case _thisArrow.style
                    Case arrowStyles.Straight
                        Dim path As ISegmentCollection = New Polyline
                        Dim firstSegment As ILine = New ESRI.ArcGIS.Geometry.Line
                        firstSegment.FromPoint = _firstPoint '.PutCoords(_firstPoint.X, _firstPoint.Y)
                        firstSegment.ToPoint = middlePoint '.PutCoords(middlePoint.X, middlePoint.Y)
                        path.AddSegment(DirectCast(firstSegment, ISegment), missing, missing)
                        thePolyline = DirectCast(path, IPolyline)
                    Case arrowStyles.Leader
                        Dim path As ISegmentCollection = New Polyline
                        Dim firstSegment As ILine = New ESRI.ArcGIS.Geometry.Line
                        Dim lastSegment As ILine = New ESRI.ArcGIS.Geometry.Line
                        firstSegment.FromPoint = _firstPoint '.PutCoords(_firstPoint.X, _firstPoint.Y)
                        firstSegment.ToPoint = middlePoint '.PutCoords(middlePoint.X, middlePoint.Y)
                        lastSegment.FromPoint = middlePoint
                        lastSegment.ToPoint = endPoint

                        path.AddSegment(DirectCast(firstSegment, ISegment), missing, missing)
                        If _pointNumber = 3 Then
                            path.AddSegment(DirectCast(lastSegment, ISegment), missing, missing)
                        End If
                        thePolyline = DirectCast(path, IPolyline)
                        firstSegment = Nothing
                        lastSegment = Nothing
                        path = Nothing
                    Case arrowStyles.Zigzag
                        segpoints(0).PutCoords(0, 0)
                        segpoints(1).PutCoords(_zigzagPosition, 0)
                        segpoints(2).PutCoords(_zigzagPosition, _zigzagWidth * offset)
                        segpoints(3).PutCoords(20, _zigzagWidth * offset)

                        Dim firstSegment As ISegment = New ESRI.ArcGIS.Geometry.Line
                        Dim lastSegment As ISegment = New ESRI.ArcGIS.Geometry.Line
                        firstSegment.FromPoint = segpoints(0)
                        firstSegment.ToPoint = segpoints(1)
                        lastSegment.FromPoint = segpoints(2)
                        lastSegment.ToPoint = segpoints(3)

                        Dim bSpline As IBezierCurveGEN = New BezierCurve
                        Dim points(3) As IPoint
                        For count = 0 To 3
                            points(count) = New ESRI.ArcGIS.Geometry.Point
                        Next

                        points(0).PutCoords(segpoints(1).X, segpoints(1).Y)
                        points(1).PutCoords(segpoints(1).X + _zigzagCurve, segpoints(1).Y)
                        points(2).PutCoords(segpoints(2).X - _zigzagCurve, segpoints(2).Y)
                        points(3).PutCoords(segpoints(2).X, segpoints(2).Y)

                        bSpline.PutCoords(points)

                        Dim path As ISegmentCollection = New Polyline
                        path.AddSegment(firstSegment)
                        path.AddSegment(DirectCast(bSpline, ISegment))
                        path.AddSegment(lastSegment)

                        thePolyline = DirectCast(path, IPolyline)

                        'use a new line to get the angle of the vector between the points
                        Dim polyVector As ILine = New Line
                        polyVector.FromPoint = thePolyline.FromPoint
                        polyVector.ToPoint = thePolyline.ToPoint
                        Dim polyAngle As Double
                        polyAngle = polyVector.Angle

                        Dim vector As ILine = New Line
                        vector.FromPoint = _firstPoint
                        vector.ToPoint = endPoint
                        Dim vectorAngle As Double = vector.Angle - polyAngle

                        Dim scale As Double
                        scale = vector.Length / polyVector.Length

                        'Transform the polylines by moving, rotating and scaling them to the _
                        'proper positions
                        Dim transform As ITransform2D = DirectCast(thePolyline, ITransform2D)
                        transform.Move(_firstPoint.X, _firstPoint.Y)
                        transform.Rotate(_firstPoint, vectorAngle)
                        transform.Scale(_firstPoint, scale, scale)
                    Case arrowStyles.Freeform
                        Dim points(_pointNumber - 1) As IPoint
                        ReDim Preserve _freeformPoints(_pointNumber - 1)
                        _freeformPoints(_pointNumber - 1) = New ESRI.ArcGIS.Geometry.Point
                        _freeformPoints(_pointNumber - 1).PutCoords(endPoint.X, endPoint.Y)

                        For count = 0 To UBound(points)
                            points(count) = New ESRI.ArcGIS.Geometry.Point
                            points(count).PutCoords(_freeformPoints(count).X, _freeformPoints(count).Y)
                        Next

                        'points(UBound(points)).PutCoords(endPoint.X, endPoint.Y)

                        Dim bezierFeedback As INewBezierCurveFeedback
                        Try
                            bezierFeedback = New NewBezierCurveFeedback
                            bezierFeedback.Start(points(0))
                            For count = 1 To UBound(points)
                                bezierFeedback.AddPoint(points(count))
                            Next
                            'bezierFeedback.AddPoint(points(UBound(points)))
                            Dim geoMetry As IGeometry
                            geoMetry = bezierFeedback.Stop
                            thePolyline = DirectCast(geoMetry, IPolyline)
                        Catch ex As Exception
                            MsgBox(ex.Message)
                        End Try
                    Case arrowStyles.RoadTic
                        Dim path As ISegmentCollection = New Polyline
                        Dim firstSegment As ILine = New ESRI.ArcGIS.Geometry.Line
                        firstSegment.FromPoint = _firstPoint
                        firstSegment.ToPoint = middlePoint
                        path.AddSegment(DirectCast(firstSegment, ISegment), missing, missing)
                        thePolyline = DirectCast(path, IPolyline)
                End Select

                If _arrowheadIsSwitched Then
                    thePolyline.ReverseOrientation()
                End If
            Catch ex As Exception
                MessageBox.Show("getSingleArrowGeometry - " & ex.ToString)
            End Try
        End If
        Return thePolyline
    End Function

    ''' <summary>
    ''' Draws the arrow image to the screen
    ''' </summary>
    ''' <param name="arrow">Input polyline</param>
    ''' <remarks></remarks>
    Friend Sub drawArrowImage(ByVal arrow As IPolyline, ByVal arrowNumber As Integer)
        Try
            Dim graphicsContainer As IGraphicsContainer = _mxDoc.ActiveView.GraphicsContainer
            Dim symbol As ISymbol

            Dim theTemplate As IEditTemplate = _editor.CurrentTemplate
            Dim editLayer As IEditLayers = DirectCast(_editor, IEditLayers)
            Dim theLayer As ILayer = theTemplate.Layer '_editor.CurrentTemplate.Layer
            Dim theFeatureLayer As IFeatureLayer = DirectCast(theLayer, IFeatureLayer)

            Dim uvr As IUniqueValueRenderer
            Dim geoLayer As IGeoFeatureLayer = CType(theFeatureLayer, IGeoFeatureLayer)
            uvr = DirectCast(geoLayer.Renderer, IUniqueValueRenderer)
            uvr.FieldCount = 1
            uvr.Field(0) = "linetype"

            symbol = uvr.DefaultSymbol

            Dim count As Integer
            For count = 0 To uvr.ValueCount
                If CInt(uvr.Value(count)) = CInt(_editor.CurrentTemplate.DefaultValue("linetype")) Then
                    symbol = DirectCast(uvr.Symbol(uvr.Value(count)), ISymbol)
                    Exit For
                End If
            Next

            Dim lineElement As ILineElement = New LineElement
            lineElement.Symbol = DirectCast(symbol, ILineSymbol)
            'Dim element As ESRI.ArcGIS.Carto.IElement
            _graphicElement(arrowNumber) = DirectCast(lineElement, IElement)
            _graphicElement(arrowNumber).Geometry = DirectCast(arrow, IGeometry)

            graphicsContainer.AddElement(_graphicElement(arrowNumber), 0)
            _mxDoc.ActiveView.PartialRefresh(esriViewDrawPhase.esriViewGraphics, Nothing, Nothing)
            '_editor.InvertAgent(_editor.Location, _mxDoc.ActivatedView.ScreenDisplay.hDC)
        Catch ex As Exception
            MessageBox.Show("drawArrowImage - " & ex.ToString)
        End Try
    End Sub

    ''' <summary>
    ''' KeyDown events 
    ''' </summary>
    ''' <param name="keyCode">the key code</param>
    ''' <param name="isShifted">the shift key status</param>
    ''' <remarks></remarks>
    Public Sub keyCommands(ByVal keyCode As Integer, ByVal isShifted As Boolean)
        Try
            If keyCode = Windows.Forms.Keys.Space And Not isShifted Then
                _arrowScale = _arrowScale * 0.9
                showLineFeedback(_lastPoint)
            ElseIf keyCode = Windows.Forms.Keys.Space And isShifted Then
                _arrowScale = _arrowScale * 1.2
                showLineFeedback(_lastPoint)
            ElseIf keyCode = Windows.Forms.Keys.F Then
                _flipArrows = Not _flipArrows
                If Not _pointNumber = 1 Then
                    showLineFeedback(_lastPoint)
                End If
            ElseIf keyCode = Windows.Forms.Keys.U Then
                _angleIsSet = Not _angleIsSet
            ElseIf keyCode = Windows.Forms.Keys.Escape Then
                If _thisArrow.style = arrowStyles.Freeform And _pointNumber > 2 Then
                    _pointNumber -= 1
                Else
                    clearAll()
                End If
            ElseIf keyCode = Windows.Forms.Keys.S Then
                _arrowheadIsSwitched = Not _arrowheadIsSwitched
                showLineFeedback(_lastPoint)
            End If
        Catch ex As Exception
            MessageBox.Show("keyCommands - " & ex.ToString)
        End Try
    End Sub

    ''' <summary>
    ''' Activates the Select Features Tool to unselect an arrow tool due to 
    ''' being in the layout view
    ''' </summary>
    ''' <remarks></remarks>
    Friend Sub setDefaultTool()
        Try
            Dim uid As New UID
            uid.Value = "esriArcMapUI.selectFeaturesTool"
            'Dim application As ESRI.ArcGIS.Framework.IApplication = DirectCast(My.ArcMap.Application, ESRI.ArcGIS.Framework.IApplication)
            'application.CurrentTool = application.Document.CommandBars.Find(uid)
            My.ArcMap.Application.CurrentTool = My.ArcMap.Application.Document.CommandBars.Find(uid)
        Catch ex As Exception
            MessageBox.Show("setDefaultTool - " & ex.ToString)
        End Try
    End Sub

    ''' <summary>
    ''' Places a freeform arrow when the user double-clicks or selects finish from
    ''' the context menu
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub placeFreeformArrow()
        Try
            placeArrows(_lastPoint)
            _pointNumber = 1
            ReDim _freeformPoints(0)
        Catch ex As Exception
            MessageBox.Show("placeFreeformArrow - " & ex.ToString)
        End Try
    End Sub

    ''' <summary>
    ''' Checks to see that the map in is data view. If not gives the option to change it.
    ''' </summary>
    ''' <returns>If the return is false the tool is cancelled</returns>
    ''' <remarks></remarks>
    Friend Function checkDataView() As Boolean
        Dim retVal As Boolean = True
        Try
            Dim mxDoc As IMxDocument = DirectCast(My.ArcMap.Application.Document, IMxDocument)
            If mxDoc.ActiveView Is mxDoc.PageLayout Then
                If MsgBox("You must be in Data View to use this tool. " & _
                    "Click OK to change to Data View.", MsgBoxStyle.OkCancel) = MsgBoxResult.Ok Then
                    mxDoc.ActiveView = DirectCast(mxDoc.FocusMap, IActiveView)
                Else
                    retVal = False
                End If
            End If
        Catch ex As Exception
            MessageBox.Show("checkDataView - " & ex.ToString)
        End Try
        Return retVal
    End Function

    ''' <summary>
    ''' Clears the graphics container
    ''' </summary>
    ''' <remarks></remarks>
    Friend Sub clearAll()
        Try
            Dim theGraphicsContainer As IGraphicsContainer = _mxDoc.ActiveView.GraphicsContainer
            _pointNumber = 1

            Try
                theGraphicsContainer.DeleteElement(_graphicElement(0))
                theGraphicsContainer.DeleteElement(_graphicElement(1))
            Catch
                'Debug.Print("No graphic element to delete")
            End Try

            _mxDoc.ActiveView.PartialRefresh(esriViewDrawPhase.esriViewGraphics, Nothing, Nothing)
        Catch ex As Exception
            MessageBox.Show("clearAll - " & ex.ToString)
        End Try
    End Sub

    Friend Function checkXML() As Boolean
        Dim xmlFile As String = _installationFolder & "\ArrowSettings.xml"
        If File.Exists(xmlFile) Then
            Return True
        Else
            Dim xmlBackup As String = xmlFile.Replace(".xml", ".bak")
            If File.Exists(xmlBackup) Then
                FileCopy(xmlBackup, xmlFile)
                Return True
            Else
                MessageBox.Show("The settings file, " & xmlFile & " , does not exist.", "File not found",
                            MessageBoxButtons.OK)
                clearAll()
                Return False
            End If
        End If
    End Function

    ''' <summary>
    ''' Reads the XML settings file
    ''' </summary>
    ''' <param name="arrowType">The arrow geometry being extracted from the file</param>
    ''' <returns>A formatted string with coordinate pairs</returns>
    ''' <remarks>The XML file is in the program installation folder</remarks>
    Private Function ReadXML(ByVal arrowType As Integer) As String
        ReadXML = ""
        Try
            Dim xmlDoc As XmlDocument = New XmlDocument
            Dim xmlNodeList As XmlNodeList
            Dim xmlNode As XmlNode

            If checkXML() = True Then

                xmlDoc.Load(_installationFolder & "\ArrowSettings.xml")

                Dim nodeName As String = ""

                Select Case arrowType
                    Case arrowCategories.Straight
                        nodeName = "/arrowTools/arrowDef/straight"
                    Case arrowCategories.LandHook
                        If _flipArrows Then
                            nodeName = "/arrowTools/arrowDef/landHookFlipped"
                        Else
                            nodeName = "/arrowTools/arrowDef/landHook"
                        End If
                    Case arrowCategories.NoDashes
                        nodeName = "/arrowTools/arrowDef/curved0"
                    Case arrowCategories.OneDash
                        nodeName = "/arrowTools/arrowDef/curved1"
                    Case arrowCategories.TwoDashes
                        nodeName = "/arrowTools/arrowDef/curved2"
                    Case arrowCategories.ThreeDashes
                        nodeName = "/arrowTools/arrowDef/curved3"
                    Case arrowCategories.FourDashes
                        nodeName = "/arrowTools/arrowDef/curved4"
                    Case arrowCategories.RoadTic
                        nodeName = "/arrowTools/arrowDef/roadTic"
                    Case Else
                        Return ""
                        Exit Function
                End Select

                'Loop through the nodes
                Dim count As Integer

                xmlNodeList = xmlDoc.SelectNodes(nodeName)
                Dim xmlChild As XmlNode
                For Each xmlNode In xmlNodeList
                    For Each xmlChild In xmlNode
                        For count = 0 To xmlChild.Attributes.Count - 1
                            ReadXML = ReadXML & xmlChild.Attributes.Item(count).InnerText & ","
                        Next
                    Next
                Next
                ReadXML = Left(ReadXML, Len(ReadXML) - 1)
                xmlDoc = Nothing
            Else
                _pointNumber = 1
            End If
        Catch ex As Exception
            MessageBox.Show("readXML - " & ex.ToString)
        End Try
    End Function

    Friend Function checkFeatureTemplate() As Boolean
        'Dim editLayer As IEditLayers = DirectCast(_editor, IEditLayers)
        Try
            Dim theLayer As ILayer = DirectCast(_editor.CurrentTemplate.Layer, ILayer)
            Dim theFeatureLayer As IFeatureLayer = DirectCast(theLayer, IFeatureLayer)
            If theFeatureLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolyline Then
                Return True
                Exit Function
            Else
                MessageBox.Show("Please select an arrow template then reselect the tool", "Incorrect Template",
                    MessageBoxButtons.OK)
                checkFeatureTemplate = False
            End If
        Catch ex As Exception
            MessageBox.Show("Please select an arrow template then reselect the tool", "Incorrect Template",
                 MessageBoxButtons.OK)
            checkFeatureTemplate = False
        Finally
        End Try
    End Function
End Module
