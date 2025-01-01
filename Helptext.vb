Public Class Helptext
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.


    End Sub

    Protected Overrides Sub Finalize()
        Me.Close()
    End Sub

    Private Sub Helptext_Leave(sender As Object, e As EventArgs) Handles MyBase.Leave
        Me.Close()
    End Sub
End Class