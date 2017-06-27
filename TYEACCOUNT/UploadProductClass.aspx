<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Site1.Master" CodeBehind="UploadProductClass.aspx.vb" Inherits="TYEACCOUNT.WebForm4" %>
<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <h2>Upload Product Class</h2>
    <div class="download">
    
    <asp:ImageButton ImageUrl="~/images/excel-icon.png"  ID="btnDownload" runat="server" Height="36px" Width="89px" ImageAlign="Middle" OnClick="btnDownload_Click" />
    </div>
    <table>
        <tr><td>
            <asp:FileUpload ID="FileUpload1" runat="server" Width="468px" />
        </td></tr>
        <tr>
            <td>
                <asp:Button ID="btnImportClass" Text="Import Class" runat="server" Width="119px" />
                <br />
                <br />
                
            </td>
        </tr>
    </table>
    <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" CellPadding="4" EnableModelValidation="True" ForeColor="#333333" GridLines="None">
        <AlternatingRowStyle BackColor="White" />
        <Columns>
            <asp:BoundField DataField="CLCLCD" HeaderText="CLASS" />
            <asp:BoundField DataField="CLMTMN" HeaderText="MATERIAL" />
            <asp:BoundField DataField="CLMTTP" HeaderText="TYPE" />
        </Columns>
                <EditRowStyle BackColor="#7C6F57" />
        <FooterStyle BackColor="#1C5E55" Font-Bold="True" ForeColor="White" />
        <HeaderStyle BackColor="#1C5E55" Font-Bold="True" ForeColor="White" />
        <PagerStyle BackColor="#666666" ForeColor="White" HorizontalAlign="Center" />
        <RowStyle BackColor="#E3EAEB" />
        <SelectedRowStyle BackColor="#C5BBAF" Font-Bold="True" ForeColor="#333333" />
                </asp:GridView>
</asp:Content>
