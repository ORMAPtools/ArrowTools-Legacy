Public Class referenceScale
    Inherits ESRI.ArcGIS.Desktop.AddIns.ComboBox
    Private isShown As Boolean = False

    Public Sub New()
        With Me
            .Add("0")
            .Add("10")
            .Add("20")
            .Add("30")
            .Add("40")
            .Add("50")
            .Add("100")
            .Add("200")
            .Add("400")
            .Add("800")
            .Add("1000")
            .Add("2000")
        End With
    End Sub

    Protected Overrides Sub OnSelChange(cookie As Integer)
        MyBase.OnSelChange(cookie)
        Try
            My.ArcMap.Document.FocusMap.ReferenceScale = CDbl(Me.GetItem(cookie).Caption) * 12
            My.ArcMap.Document.ActiveView.Refresh()
        Catch

        End Try
    End Sub

    Protected Overrides Sub OnUpdate()
        MyBase.OnUpdate()
        Try
            Enabled = My.ArcMap.Application IsNot Nothing
            If Not isShown Then
                If My.ArcMap.Document.FocusMap.MapScale > 0 Then
                    Dim refScale As Double = My.ArcMap.Document.FocusMap.ReferenceScale
                    For i As Int32 = 0 To Me.items.Count - 1
                        If CDbl(Me.items(i).Caption) * 12 = refScale Then
                            Me.Select(Me.items(i).Cookie)
                            Exit For
                        End If
                    Next
                    isShown = True
                End If
            End If
        Catch ex As Exception

        End Try
    End Sub
End Class
