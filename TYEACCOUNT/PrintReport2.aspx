<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Site1.Master" CodeBehind="PrintReport2.aspx.vb" Inherits="TYEACCOUNT.WebForm7" %>

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
    <h2>Report 2 : Sale Comparison By Class </h2> 
    <table>
        <tr>
            <td>Invoice Date From : </td>
            <td>
                <asp:TextBox ID="txtDateS" runat="server"></asp:TextBox>
                <asp:RequiredFieldValidator ID="Req3" runat="server" ControlToValidate="txtDateS" ErrorMessage="*" ValidationGroup="submit"></asp:RequiredFieldValidator>
            </td>
            <td>To : </td>
            <td>
                <asp:TextBox ID="txtDateE" runat="server"></asp:TextBox>
                <asp:RequiredFieldValidator ID="Req2" runat="server" ControlToValidate="txtDateE" ErrorMessage="*" ValidationGroup="submit"></asp:RequiredFieldValidator>
            </td>
        </tr>

    </table>

    <br />
    <asp:Button ID="btnPrint" runat="server" Text="Print Report2 " ValidationGroup="submit"/>
    <br />
    <fieldset class="fsStandard" style="padding-bottom: 6px; width: 60%;">
        <legend><span>รายชื่อ Item Class ที่ยังไม่มีใน Item Master</span></legend>
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" BackColor="#DEBA84" BorderColor="#DEBA84" BorderStyle="None" BorderWidth="1px" CellPadding="3" CellSpacing="2" EnableModelValidation="True" Width="98px">
            <Columns>
                <asp:TemplateField HeaderText="Item Class">
                    <ItemTemplate>
                        <asp:Label ID="lblClass" runat="server" Text='<%# DataBinder.Eval(Container, "DataItem.SPITCL")%>' CssClass="Padding">
                        </asp:Label>
                    </ItemTemplate>
                </asp:TemplateField>
            </Columns>
            <FooterStyle BackColor="#F7DFB5" ForeColor="#8C4510" />
            <HeaderStyle BackColor="#A55129" Font-Bold="True" ForeColor="White" />
            <PagerStyle ForeColor="#8C4510" HorizontalAlign="Center" />
            <RowStyle BackColor="#FFF7E7" ForeColor="#8C4510" />
            <SelectedRowStyle BackColor="#738A9C" Font-Bold="True" ForeColor="White" />
        </asp:GridView>
    </fieldset>

</asp:Content>
