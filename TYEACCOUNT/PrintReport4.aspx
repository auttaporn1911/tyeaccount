<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Site1.Master" CodeBehind="PrintReport4.aspx.vb" Inherits="TYEACCOUNT.PrintReport4" %>
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
    <h2>Report 4 : Compare Sales Agent</h2>
    <table>
        <tr>
            <td>Invoice Date From : </td>
            <td><asp:TextBox ID="txtDateS" runat="server"></asp:TextBox></td>
            <td>To : </td>
            <td><asp:TextBox ID="txtDateE" runat="server"></asp:TextBox></td>
        </tr>
    </table>
   
    <br />
    <asp:Button ID="btnPrint" runat="server" Text="Print Report4" />
    <br /><br />
    <asp:Label ID="Label2" runat="server" Text="คำเตือน : เนื่องจากมี Class บาง Class ยังไม่เคยถูกบันทึกเข้าระบบมาก่อน จึงอาจทำให้การออกรายงานผิดพลาด กรุณานำเข้า Class ดังกล่าว ก่อน PrintReport" Style="font-size:14px; color:#ae0606" Visible="false"></asp:Label>
     <br /><asp:Label ID="Label1" runat="server" Text="รายชื่อ Class ที่ไม่มีในระบบมีดังนี้ :" Style="font-size:14px;" Visible="false"></asp:Label>
    <br />
     <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="false" BorderColor="#cccccc">
        <RowStyle BackColor="#ffcaca" ForeColor="#990000" Font-Names="Tahoma" Font-Size="12px" />
        <PagerStyle BackColor="#ffd2d2" ForeColor="#990000" HorizontalAlign="Right" Font-Names="Tahoma" Font-Size="12px"/>
        <SelectedRowStyle BackColor="#738A9C" Font-Bold="True" ForeColor="#F7F7F7" Font-Names="Tahoma" Font-Size="12px"/>
        <HeaderStyle BackColor="#8a2828" Font-Bold="True" ForeColor="#ffffff" Font-Names="Tahoma" Font-Size="14px"/>
        <AlternatingRowStyle BackColor="#F7F7F7" />
        <Columns>
            <asp:TemplateField HeaderText=" &nbsp;Class&nbsp; " HeaderStyle-Height="30px">
                <ItemTemplate>
                    <asp:Label ID="lblLotno" runat="server" Text='<%# DataBinder.Eval(Container, "DataItem.Class") %>' CssClass="Padding">
                    </asp:Label>
                </ItemTemplate>
                <ItemStyle HorizontalAlign="Center" />
            </asp:TemplateField>
            
        </Columns>
    </asp:GridView>
</asp:Content>