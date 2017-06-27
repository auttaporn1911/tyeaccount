<%@ Page Language="VB" AutoEventWireup="false" CodeFile="login.aspx.vb" Inherits="login" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Login</title>
    <link rel="stylesheet" type="text/css" href="App_Themes/Theme1/style1.css" />
</head>

<body>
    <form id="form1" runat="server">
     <div style="color: white; font-size: 22px; padding:5px;"> SALES WEB TYE SYSTEM </div>
  <div class="inset">
  <p>
    <label for="email">
        USER NAME</label>
    <asp:TextBox id="txtUsername" runat="server"/>
  </p>
  <p>
    <label for="password">PASSWORD</label>
    <asp:TextBox id="txtPassword" runat="server" TextMode="Password"/>
  </p>
  
  </div>
  <p class="p-container">
    <span></span>
    <asp:Button ID="btnLogin" Text="Login" runat="server" OnClick="btnLogin_Click" />
  </p>
        <p>
            &nbsp;</p>
    </form>
</body>
</html>
