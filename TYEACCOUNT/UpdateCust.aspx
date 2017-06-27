<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Site1.Master" CodeBehind="UpdateCust.aspx.vb" Inherits="TYEACCOUNT.WebForm14" %>
<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <h2>Customer Information</h2>
    <table>
        <tr>
            <td>

                Customer Code :

            </td>
            <td>
                <asp:TextBox ID="txtCustCode" runat="server"></asp:TextBox>
            </td>
        </tr>
        <tr>
            <td>

                Account Code :</td>
            <td>
                <asp:TextBox ID="txtAccCode" runat="server"></asp:TextBox>
            </td>
        </tr>
        <tr>
            <td>

                Customer Name :

            </td>
            <td>
                <asp:TextBox ID="txtName" runat="server" Width="274px"></asp:TextBox>
            </td>
        </tr>
        <tr>
            <td>

                Type :</td>
            <td>
                <asp:TextBox ID="txtType" runat="server"></asp:TextBox>
            </td>
        </tr>
        <tr>
            <td>

                Group :</td>
            <td>
                <asp:TextBox ID="txtGroup" runat="server"></asp:TextBox>
            </td>
        </tr>
        <tr>
            <td>

                Section :</td>
            <td>
                <asp:TextBox ID="txtSection" runat="server"></asp:TextBox>
            </td>
        </tr>
        <tr>
            <td>

                Custom :</td>
            <td>
                <asp:TextBox ID="txtCustom" runat="server"></asp:TextBox>
            </td>
        </tr>
        <tr>
            <td>

                Effective Date :</td>
            <td>
                <asp:TextBox ID="txtEffDate" runat="server"></asp:TextBox>
            </td>
        </tr>
        <tr>
            <td>

                &nbsp;</td>
            <td>
                <asp:Button ID="btnSave" runat="server" Text="Save" Width="58px" />
            </td>
        </tr>
    </table>
</asp:Content>
