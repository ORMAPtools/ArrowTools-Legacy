Imports System.Drawing
Imports System
Imports ESRI.ArcGIS.SystemUI
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Geometry
Imports ESRI.ArcGIS.Framework
Imports System.Windows.Forms
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Editor
Imports ESRI.ArcGIS.Display

Public Class DimensionArrows
    Inherits ESRI.ArcGIS.Desktop.AddIns.Tool

    Public prevPoint As IPoint = New ESRI.ArcGIS.Geometry.Point

    Public Sub New()
        Dim editorUID As New ESRI.ArcGIS.esriSystem.UID
        editorUID.Value = "esriEditor.editor"

        _editor = DirectCast(My.ArcMap.Application.FindExtensionByCLSID(editorUID), IEditor3)
        _editEvents = CType(CType(_editor, IEditEvents_Event), Editor)
        _installationFolder = System.IO.Path.GetDirectoryName(Me.GetType().Assembly.Location)
    End Sub

    Protected Overrides Sub OnUpdate()
       Enabled = (_editor.EditState = esriEditState.esriStateEditing)
        setMouseCursor()
    End Sub

    Protected Overrides Sub OnActivate()
        MyBase.OnActivate()
        Try
            clearAll()
            _pointNumber = 1
            _thisArrow.category = arrowCategories.NoDashes
            _selectNewArrows = CBool(GetSetting("OR_DOR_dimensionArrows", "default", "selectNewArrows", CStr(True)))
            MyBase.Cursor = New Windows.Forms.Cursor(Me.GetType(), "curved0.cur")
            If checkFeatureTemplate() = False Then
                clearAll()
                Exit Sub
            End If
        Catch ex As Exception
            MessageBox.Show("OnActivate - " & ex.ToString)
        End Try
    End Sub


    Protected Overrides Sub OnMouseDown(arg As ESRI.ArcGIS.Desktop.AddIns.Tool.MouseEventArgs)
        MyBase.OnMouseDown(arg)
       
        Try
            Dim snap As ISnapEnvironment = DirectCast(_editor, ISnapEnvironment)
            Dim clickPoint As New ESRI.ArcGIS.Geometry.Point
            clickPoint = CType(_mxDoc.CurrentLocation, ESRI.ArcGIS.Geometry.Point)
            snap.SnapPoint(clickPoint)

            If arg.Button = Windows.Forms.MouseButtons.Left Then
                If _pointNumber = 1 Then
                    setLineFeedback(clickPoint) '_mxDoc.CurrentLocation)
                    _pointNumber = 2
                ElseIf _pointNumber = 2 Then
                    showLineFeedback(_lastPoint)
                    _pointNumber = 3
                Else
                    setArrowOffset(clickPoint) '_mxDoc.CurrentLocation)
                    placeArrows(_lastPoint)
                    prevPoint.PutCoords(0, 0)
                    _pointNumber = 1
                End If
            ElseIf arg.Button = Windows.Forms.MouseButtons.Right Then
                CreateContextMenu()
            End If
        Catch ex As Exception
            MessageBox.Show("onMouseUp - " & ex.ToString)
        End Try
    End Sub

    Protected Overrides Sub OnMouseMove(arg As ESRI.ArcGIS.Desktop.AddIns.Tool.MouseEventArgs)
        MyBase.OnMouseMove(arg)
        Try
            checkSnap()

            Dim geoPoint As IPoint = _mxDoc.CurrentLocation
           
            If _pointNumber > 1 Then

                _angleIsSet = False

                If _pointNumber = 3 Then
                    setArrowOffset(geoPoint)
                    showLineFeedback(_lastPoint)
                ElseIf _pointNumber = 2 Then
                    showLineFeedback(geoPoint)
                End If

                prevPoint = geoPoint
            End If
        Catch ex As Exception
            MessageBox.Show("OnMouseMove - " & ex.ToString)
        End Try
    End Sub


    ''' <summary>
    ''' Sets the offset distance from an imaginary line between the two arrow points
    ''' </summary>
    ''' <param name="nearPoint"></param>
    ''' <remarks></remarks>
    Private Sub setArrowOffset(ByVal nearPoint As IPoint)
        Try
            Dim displayTransformation As IDisplayTransformation
            displayTransformation = My.ThisApplication.Display.DisplayTransformation

            Dim fromPoint As IPoint = New ESRI.ArcGIS.Geometry.Point
            fromPoint.PutCoords(_firstPoint.X, _firstPoint.Y)
            Dim toPoint As IPoint = New ESRI.ArcGIS.Geometry.Point
            toPoint.PutCoords(_lastPoint.X, _lastPoint.Y)

            'use a new line to get the angle of the vector between the points
            Dim vector As ILine = New Line

            vector.FromPoint = fromPoint
            vector.ToPoint = toPoint

            Dim rightSide As Boolean

            vector.QueryPointAndDistance(esriSegmentExtension.esriExtendAtFrom, nearPoint, False, _
                Nothing, Nothing, _arrowOffset, rightSide)
            If rightSide Then _arrowOffset = _arrowOffset * -1
        Catch ex As Exception
            MessageBox.Show("setArrowOffset - " & ex.ToString)
        End Try
    End Sub

    Protected Overrides Sub OnKeyDown(arg As ESRI.ArcGIS.Desktop.AddIns.Tool.KeyEventArgs)
        MyBase.OnKeyDown(arg)
        Try
            If arg.Shift = False Then
                Select Case arg.KeyCode
                    Case Windows.Forms.Keys.D0, Windows.Forms.Keys.NumPad0
                        _thisArrow.category = arrowCategories.NoDashes
                    Case Windows.Forms.Keys.D1, Windows.Forms.Keys.NumPad1
                        _thisArrow.category = arrowCategories.OneDash
                    Case Windows.Forms.Keys.D2, Windows.Forms.Keys.NumPad2
                        _thisArrow.category = arrowCategories.TwoDashes
                    Case Windows.Forms.Keys.D3, Windows.Forms.Keys.NumPad3
                        _thisArrow.category = arrowCategories.ThreeDashes
                    Case Windows.Forms.Keys.D4, Windows.Forms.Keys.NumPad4
                        _thisArrow.category = arrowCategories.FourDashes
                    Case Else
                        keyCommands(arg.KeyCode, arg.Shift)
                End Select
                setMouseCursor()

                If Not _pointNumber = 1 Then
                    showLineFeedback(_lastPoint)
                End If
            Else
                keyCommands(arg.KeyCode, arg.Shift)
            End If
        Catch ex As Exception
            MessageBox.Show("onKeyDown - " & ex.ToString)
        End Try
    End Sub

    Friend Sub setMouseCursor()
        Select Case _thisArrow.category
            Case arrowCategories.NoDashes
                MyBase.Cursor = New System.Windows.Forms.Cursor(Me.GetType(), "curved0.cur")
            Case arrowCategories.OneDash
                MyBase.Cursor = New System.Windows.Forms.Cursor(Me.GetType(), "curved1.cur")
            Case arrowCategories.TwoDashes
                MyBase.Cursor = New System.Windows.Forms.Cursor(Me.GetType(), "curved2.cur")
            Case arrowCategories.ThreeDashes
                MyBase.Cursor = New System.Windows.Forms.Cursor(Me.GetType(), "curved3.cur")
            Case arrowCategories.FourDashes
                MyBase.Cursor = New System.Windows.Forms.Cursor(Me.GetType(), "curved4.cur")
        End Select
    End Sub

    Private Sub CreateContextMenu()
        Try
                 Dim commandBars As ICommandBars = My.ArcMap.Application.Document.CommandBars
                Dim commandBar As ICommandBar = commandBars.Create("TemporaryContextMenu",
                    ESRI.ArcGIS.SystemUI.esriCmdBarType.esriCmdBarTypeShortcutMenu)

                Dim optionalIndex As System.Object = Type.Missing
                Dim uid As ESRI.ArcGIS.esriSystem.UID = New ESRI.ArcGIS.esriSystem.UID

                uid.Value = "ORMAP_ArrowTools_arrowContextMenu"
                uid.SubType = 0
                commandBar.Add(uid, optionalIndex)

                With _menuitems
                    .Clear()
                If _pointNumber > 1 Then
                    .Add(SHORTER)
                    .Add(LONGER)
                    .Add(SWITCH)
                    .Add(CANCEL)
                    .Add(SEPARATOR)
                    .Add(SCALE10)
                    .Add(SCALE20)
                    .Add(SCALE30)
                    .Add(SCALE40)
                    .Add(SCALE50)
                    .Add(SCALE100)
                    .Add(SCALE200)
                    .Add(SCALE400)
                    .Add(SCALE800)
                    .Add(SCALE1000)
                    .Add(SCALE2000)
                    .Add(SEPARATOR)
                    .Add(SELECTION)
                    .Add(HELP)
                Else
                    .Add(SELECTION)
                    .Add(HELP)
                End If
                End With

                commandBar.Popup()

        Catch ex As Exception
            MessageBox.Show("createContextMenu - " & ex.ToString)
        End Try
    End Sub

    Protected Overrides Function OnDeactivate() As Boolean
        clearAll()
        Return True
    End Function
End Class
