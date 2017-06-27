<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Site1.Master" CodeBehind="PrintReport3.aspx.vb" Inherits="TYEACCOUNT.WebForm12" %>
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
    <h2>Report 3 : SALE COMPARISON BY CUSTOMER</h2>
    <table>
        <tr>
            <td>Invoice Date From : </td>
            <td><asp:TextBox ID="txtDateS" runat="server"></asp:TextBox>
                <asp:RequiredFieldValidator ID="Req3" runat="server" ControlToValidate="txtDateS" ErrorMessage="*" ValidationGroup="submit"></asp:RequiredFieldValidator>
             </td>
            <td>To : </td>
            <td><asp:TextBox ID="txtDateE" runat="server"></asp:TextBox>
                <asp:RequiredFieldValidator ID="Req2" runat="server" ControlToValidate="txtDateE" ErrorMessage="*" ValidationGroup="submit"></asp:RequiredFieldValidator>
             </td>
        </tr>
    </table>
    <br />
    <asp:Button ID="btnPrint" runat="server" Text="Print Report3" ValidationGroup="submit" />
     <br />
     <br />
    <fieldset class="fsStandard" style="padding-bottom: 6px; width: 60%;">
        <legend>รายชื่อลูกค้าที่ยังไม่มีใน File Mapping</legend>
        <asp:GridView ID="gvCheck" runat="server" BackColor="#DEBA84" BorderColor="#DEBA84" BorderStyle="None" BorderWidth="1px" CellPadding="3" CellSpacing="2" EnableModelValidation="True">
            <FooterStyle BackColor="#F7DFB5" ForeColor="#8C4510" />
            <HeaderStyle BackColor="#A55129" Font-Bold="True" ForeColor="White" />
            <PagerStyle ForeColor="#8C4510" HorizontalAlign="Center" />
            <RowStyle BackColor="#FFF7E7" ForeColor="#8C4510" />
            <SelectedRowStyle BackColor="#738A9C" Font-Bold="True" ForeColor="White" />
        </asp:GridView>
    </fieldset>
     <br />
     <br />
</asp:Content>
