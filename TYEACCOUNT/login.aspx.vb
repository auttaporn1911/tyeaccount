Imports System.Data
Imports system.Data.OleDb
Partial Class login
    Inherits System.Web.UI.Page

    Protected Sub btnLogin_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnLogin.Click
        Dim username As String
        Dim pwd As String

        username = Trim(txtUsername.Text)
        pwd = Trim(txtPassword.Text)
        If username = "" Or pwd = "" Then
            MessageBox("Please input username or password")
            Exit Sub
        End If
        'คิวรี่ข้อมูลแบบปกติ*************************)**********************
        Dim sqlUserName As String

        sqlUserName = "SELECT  users.defgrp,USERS.UserID, USERS.EmpCode, USERS.EmpName, USERS.EMELNAME,users.DEPARTMENT,users.DEFCOM, "
        sqlUserName += "users.ASUID,WEBSalePermis.MenuID, "
        sqlUserName += "WEBSaleMenu.MName "
        sqlUserName += "FROM   SYSTEMNAME "
        sqlUserName += "INNER JOIN WEBSaleMenu ON SYSTEMNAME.SCODE = WEBSaleMenu.SYSNAME "
        sqlUserName += "INNER Join WEBSalePermis ON WEBSaleMenu.MID = WEBSalePermis.MenuID "
        sqlUserName += "INNER JOIN USERS ON WEBSalePermis.UID = USERS.UserID "
        sqlUserName += "WHERE SYSTEMNAME.SCODE = 'SALEWEB' "
        sqlUserName += "AND (Users.UserID = ? ) "
        sqlUserName += "AND (Users.Password = ? ) "

        '  response.write(sqlUserName)
        Dim Conn As New OleDbConnection(Classconn.strConnSql)
        Dim UserDA As New OleDbDataAdapter

        Dim selectCMD As New OleDbCommand(sqlUserName, Conn)
        UserDA.SelectCommand = selectCMD
        ' Add parameters and set values.
        selectCMD.Parameters.Add("@UserID", OleDbType.VarChar, 15).Value = username
        selectCMD.Parameters.Add("@Password", OleDbType.VarChar, 15).Value = pwd
        'selectCMD.Parameters.Add("@ASUID", OleDbType.VarChar, 10).Value = AS400ID

        Dim UserDT As New DataTable
        UserDA.Fill(UserDT)

        UserDA.SelectCommand.Connection.Close()

        Session.Timeout = 20
        If UserDT.Rows.Count < 1 Then
            'Session("userid") = Session("userid")
            'Session("ULevel") = Session("uid")
            'Session("UserName") = ""
            'Session("asuid") = ""
            'Session("department") = ""
            'Session("Company") = ""
            'Session("DefGrp") = ""
            'Response.Redirect("login.aspx")
            'Response.Redirect("Default.aspx")
            MessageBox("Username or Password is not valid.")

        Else
            Session("userid") = username
            Session("UserName") = UserDT.Rows(0).Item("EmpName").ToString
            Session("userdesp") = UserDT.Rows(0).Item("UserID").ToString & "  " & UserDT.Rows(0).Item("EmpName").ToString
            Session("DTMenu") = UserDT
            Session("asuid") = UserDT.Rows(0).Item("ASUID").ToString
            Session("department") = UserDT.Rows(0).Item("Department").ToString
            Session("Company") = UserDT.Rows(0).Item("DEFCOM").ToString
            Session("DefGrp") = UserDT.Rows(0).Item("DEFGRP").ToString


            Response.Redirect("Default.aspx")

        End If
    End Sub
    Public Sub MessageBox(ByVal message As String)
        ClientScript.RegisterStartupScript(Me.GetType(), Guid.NewGuid().ToString(), String.Format("alert('{0}');", message), True)
    End Sub


End Class
