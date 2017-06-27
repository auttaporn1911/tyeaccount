<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Site1.Master" CodeBehind="PrintReport1.aspx.vb" Inherits="TYEACCOUNT.WebForm5" %>
<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
     <script>
         $(document).ready(function () {
             $("#<%=txtDateS.ClientID%>").datepicker({
                dateFormat: 'yymmdd'
             })
             $("#<%=txtDateE.ClientID%>").datepicker({
                 dateFormat: 'yymmdd'
             })
            });

    </script>
    <h2>Report 1 : Sales Break Down</h2>
    <table>
        <tr>
            <td>Invoice Date From : </td>
            <td><asp:TextBox ID="txtDateS" runat="server"></asp:TextBox>
                <asp:RequiredFieldValidator ID="req3" runat="server" ControlToValidate="txtDateS" ErrorMessage="*" ValidationGroup="submit"></asp:RequiredFieldValidator>
            </td>
            <td>To : </td>
            <td><asp:TextBox ID="txtDateE" runat="server"></asp:TextBox>
                <asp:RequiredFieldValidator ID="req2" runat="server" ControlToValidate="txtDateE" ErrorMessage="*" ValidationGroup="submit"></asp:RequiredFieldValidator>
            </td>
        </tr>
    </table>
    <br />
    <asp:Button ID="btnPrint" runat="server" Text="Print Report1" ValidationGroup="submit" />
</asp:Content>
