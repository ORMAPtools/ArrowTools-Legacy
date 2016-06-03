Imports System.Windows.Forms
Imports ESRI.ArcGIS.Editor
Imports ESRI.ArcGIS.Geometry
Imports ESRI.ArcGIS.Framework

Public Class singleArrow
    Inherits ESRI.ArcGIS.Desktop.AddIns.Tool

    Private secondPoint As IPoint = New Point
    Private prevPoint As IPoint = New Point

    Public Sub New()
        Dim editorUID As New ESRI.ArcGIS.esriSystem.UID
        editorUID.Value = "esriEditor.editor"

        _editor = DirectCast(My.ArcMap.Application.FindExtensionByCLSID(editorUID), IEditor3)
        _editEvents = CType(CType(_editor, IEditEvents_Event), Editor)
        _installationFolder = System.IO.Path.GetDirectoryName(Me.GetType().Assembly.Location)
    End Sub

    Protected Overrides Sub OnKeyDown(arg As ESRI.ArcGIS.Desktop.AddIns.Tool.KeyEventArgs)
        MyBase.OnKeyDown(arg)
        keyCommands(arg.KeyCode, arg.Shift)
    End Sub

    Protected Overrides Sub OnActivate()
        MyBase.OnActivate()
        Try
            If checkFeatureTemplate() = False Then
                clearAll()
                Exit Sub
            End If

            clearAll()
            _pointNumber = 1

            'Get the default arrow category and style from the registry
            _thisArrow.category = CInt(GetSetting("OR_DOR_dimensionArrows", "default", "category", _
                CStr(arrowCategories.SingleArrow)))
            _thisArrow.style = CInt(GetSetting("OR_DOR_dimensionArrows", "default", "style", _
                CStr(arrowStyles.Straight)))
            _selectNewArrows = CBool(GetSetting("OR_DOR_dimensionArrows", "default", "selectNewArrows", CStr(True)))
           setCursor()
            _angleIsSet = True
        Catch ex As Exception
            MessageBox.Show("OnActivate - " & ex.ToString)
        End Try
    End Sub

    Protected Overrides Sub OnMouseDown(arg As ESRI.ArcGIS.Desktop.AddIns.Tool.MouseEventArgs)
        MyBase.OnMouseDown(arg)
        Try

            Dim clickPoint As New ESRI.ArcGIS.Geometry.Point
            clickPoint.PutCoords(_mxDoc.CurrentLocation.X, _mxDoc.CurrentLocation.Y)

            Dim snap As ISnapEnvironment = DirectCast(_editor, ISnapEnvironment)
            snap.SnapPoint(clickPoint) '_mxDoc.CurrentLocation)
            checkSnap()
            '_editor.InvertAgent(prevPoint, _mxDoc.ActivatedView.ScreenDisplay.hDC)
            '_editor.InvertAgent(clickPoint, _mxDoc.ActivatedView.ScreenDisplay.hDC)

            If arg.Button = Windows.Forms.MouseButtons.Left Then
                If _pointNumber = 1 Then
                    prevPoint.PutCoords(clickPoint.X, clickPoint.Y)
                    setLineFeedback(clickPoint)
                    _pointNumber = 2
                ElseIf _thisArrow.style = arrowStyles.Leader Then
                    If _pointNumber = 2 Then
                        secondPoint.PutCoords(_snapPoint.X, _snapPoint.Y) ' clickPoint.X, clickPoint.Y)
                        showLineFeedback(clickPoint, secondPoint)
                        _pointNumber = 3
                    Else
                        placeArrows(_lastPoint, secondPoint)
                        _pointNumber = 1
                        secondPoint = New ESRI.ArcGIS.Geometry.Point
                    End If
                ElseIf _thisArrow.style = arrowStyles.Freeform Then
                    ReDim Preserve _freeformPoints(_pointNumber - 1)
                    _freeformPoints(_pointNumber - 1) = New ESRI.ArcGIS.Geometry.Point
                    _freeformPoints(_pointNumber - 1).PutCoords(clickPoint.X, clickPoint.Y)
                    showLineFeedback(clickPoint)
                    _pointNumber = _pointNumber + 1
                Else
                    placeArrows(_snapPoint) '_lastPoint)
                    _pointNumber = 1
                    'secondPoint = New ESRI.ArcGIS.Geometry.Point
                End If
            ElseIf arg.Button = Windows.Forms.MouseButtons.Right Then
                CreateContextMenu()
            End If
        Catch ex As Exception
            MessageBox.Show("onMouseDown - " & ex.ToString)
        End Try
    End Sub

    Protected Overrides Sub OnMouseMove(arg As ESRI.ArcGIS.Desktop.AddIns.Tool.MouseEventArgs)
        MyBase.OnMouseMove(arg)
        Try
            Dim geoPoint As IPoint = New ESRI.ArcGIS.Geometry.Point
            
            checkSnap()
            geoPoint.PutCoords(_mxDoc.CurrentLocation.X, _mxDoc.CurrentLocation.Y)

            If _thisArrow.style = arrowStyles.Freeform Then
                If _pointNumber > 1 Then
                    showLineFeedback(geoPoint)
                End If
            ElseIf _pointNumber = 2 Then
                '_angleIsSet = False
                showLineFeedback(geoPoint)
            ElseIf _pointNumber = 3 Then
                _arrowAngle = 0
                showLineFeedback(geoPoint, secondPoint)
            ElseIf _pointNumber > 3 Then
                showLineFeedback(geoPoint)
            End If
            prevPoint = geoPoint
        Catch ex As Exception
            MessageBox.Show("OnMouseMove - " & ex.ToString)
        End Try
    End Sub

    Protected Overrides Sub OnDoubleClick()
        MyBase.OnDoubleClick()
        If _thisArrow.style = arrowStyles.Freeform Then
            'remove the last point that was created with the mouse down event 
            ' that was triggered when double-clicking
            _pointNumber = _pointNumber - 1
            placeFreeformArrow()
        End If
    End Sub

    Friend Sub setCursor()
        'Select the cursor image as an indicator of the category of arrow being placed
        Dim cursorName As String = ""
        Select Case _thisArrow.style
            Case arrowStyles.Straight
                cursorName = "SingleStraight.cur"
            Case arrowStyles.Leader
                cursorName = "SingleLeader.cur"
            Case arrowStyles.Zigzag
                cursorName = "SingleZigzag.cur"
            Case arrowStyles.Freeform
                cursorName = "SingleFreeform.cur"
        End Select
        MyBase.Cursor = New System.Windows.Forms.Cursor(Me.GetType(), cursorName)
    End Sub

    Protected Overrides Sub OnUpdate()
        Enabled = (_editor.EditState = esriEditState.esriStateEditing)
        setCursor()
    End Sub

    Protected Overrides Function OnDeactivate() As Boolean
        clearAll()
        Return True
    End Function

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
                    Select Case _thisArrow.style
                        Case arrowStyles.Straight
                            .Add(STYLE_LEADER)
                            .Add(STYLE_ZIGZAG)
                            .Add(STYLE_FREEFORM)
                            .Add(SEPARATOR)
                            .Add(SWITCH)
                            .Add(CANCEL)
                            .Add(SEPARATOR)
                            .Add(SELECTION)
                            .Add(HELP)
                        Case arrowStyles.Freeform
                            .Add(FINISH)
                            .Add(SEPARATOR)
                            .Add(STYLE_STRAIGHT)
                            .Add(STYLE_LEADER)
                            .Add(STYLE_ZIGZAG)
                            .Add(SEPARATOR)
                            .Add(SWITCH)
                            .Add(CANCEL)
                            .Add(SEPARATOR)
                            .Add(SELECTION)
                            .Add(HELP)
                        Case arrowStyles.Leader
                            .Add(STYLE_STRAIGHT)
                            .Add(STYLE_ZIGZAG)
                            .Add(STYLE_FREEFORM)
                            .Add(SEPARATOR)
                            .Add(UNLOCK)
                            .Add(SWITCH)
                            .Add(CANCEL)
                            .Add(SEPARATOR)
                            .Add(SELECTION)
                            .Add(HELP)
                        Case arrowStyles.Zigzag
                            .Add(STYLE_STRAIGHT)
                            .Add(STYLE_LEADER)
                            .Add(STYLE_FREEFORM)
                            .Add(SEPARATOR)
                            .Add(TOPOINT)
                            .Add(TOEND)
                            .Add(NARROWER)
                            .Add(WIDER)
                            .Add(CURVELESS)
                            .Add(CURVEMORE)
                            .Add(FLIP)
                            .Add(SWITCH)
                            .Add(CANCEL)
                            .Add(SAVEDEFAULT)
                            .Add(SEPARATOR)
                            .Add(SELECTION)
                            .Add(HELP)
                    End Select
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
End Class
