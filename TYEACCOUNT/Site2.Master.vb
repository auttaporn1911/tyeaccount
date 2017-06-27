Public Class Site2
    Inherits System.Web.UI.MasterPage

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Page.IsPostBack = False Then
            'If Session("userid") Is Nothing Then
            '    Response.Redirect("UserLogin.aspx")
            'End If
        End If
    End Sub

    Protected Sub lnLogout_Click(sender As Object, e As EventArgs) Handles lnLogout.Click
        Session("userid") = Nothing
        Response.Redirect("UserLogin.aspx")
    End Sub
End Class