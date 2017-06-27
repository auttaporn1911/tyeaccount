<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Site1.Master" CodeBehind="WebForm1.aspx.vb" Inherits="TYEACCOUNT.WebForm1" %>
<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <h2>Report 1</h2>
    <table>
        <tr>
            <td>Date :</td>
            <td><asp:TextBox ID="txtDateStart" runat="server"></asp:TextBox></td>
        </tr>
        <tr>
            <td></td>
            <td><asp:Button ID="btnPrint" runat="server" Text="Print" Width="57px" /></td>
        </tr>
    </table>
</asp:Content>
