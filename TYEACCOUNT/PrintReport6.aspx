<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Site1.Master" CodeBehind="PrintReport6.aspx.vb" Inherits="TYEACCOUNT.WebForm11" %>

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
    <h2>Report 6 : CANCEL/ RETURN/DISCOUNT</h2>
    <table>
        <tr>
            <td>Invoice Date From : </td>
            <td>
                <asp:TextBox ID="txtDateS" runat="server"></asp:TextBox></td>
            <td>To : </td>
            <td>
                <asp:TextBox ID="txtDateE" runat="server"></asp:TextBox></td>
        </tr>
        
    </table>
    <br />
    <asp:Button ID="btnPrint" runat="server" Text="Print Report6" />
</asp:Content>
