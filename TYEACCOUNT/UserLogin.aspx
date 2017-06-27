<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Site2.Master" CodeBehind="UserLogin.aspx.vb" Inherits="TYEACCOUNT.WebForm13" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div class="inset">
        <p>
            <label for="email">
                USER NAME</label>
            <asp:TextBox ID="txtUsername" runat="server" />
        </p>
        <p>
            <label for="password">PASSWORD</label>
            <asp:TextBox ID="txtPassword" runat="server" TextMode="Password" />
        </p>

    </div>
    <p class="p-container">
        <span></span>
        <asp:Button ID="btnLogin" Text="Login" runat="server" OnClick="btnLogin_Click" />
    </p>
    <p>
        &nbsp;
    </p>
</asp:Content>
