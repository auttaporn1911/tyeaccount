<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Site1.Master" CodeBehind="UploadMapping.aspx.vb" Inherits="TYEACCOUNT.WebForm3" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <h2>Upload Mapping Code</h2>
    <div class="download">

        <asp:ImageButton ImageUrl="~/images/excel-icon.png" ID="btnDownload" runat="server" Height="36px" Width="97px" ImageAlign="Middle" OnClick="btnDownload_Click" />
    </div>
    <table>
        <tr>
            <td>
                <asp:FileUpload ID="FileUpload1" runat="server" Width="468px" />
            </td>
        </tr>
        <tr>
            <td>
                <asp:Button ID="btnImportCust" Text="Import Customer" runat="server" Width="119px" Enabled="False" />
                <br />
                <br />

                <asp:ImageButton ID="ImageButton1" runat="server" ImageUrl="~/images/btn-add.png" Height="31px" Width="123px" />
                <br />

            </td>
        </tr>
    </table>
    <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" CellPadding="4" EnableModelValidation="True" ForeColor="#333333" GridLines="None">
        <AlternatingRowStyle BackColor="White" />
        <columns>
            
        <asp:BoundField DataField="CSCSCD" HtmlEncode="False" DataFormatString="<a target='_parent' href='UpdateCust.aspx?cust={0}&state=update'>{0}</a>"
            HeaderText="Customer Code" />
        <asp:BoundField DataField="CSACCD" HeaderText="Account Code" />
        <asp:BoundField DataField="CSCSNM" HeaderText="Customer Name" />
        <asp:BoundField DataField="CSTYPE" HeaderText="Type" />
        <asp:BoundField DataField="CSCUTM" HeaderText="Custom" />
        <asp:BoundField DataField="CSCSGP" HeaderText="Group" />
        <asp:BoundField DataField="CSSECT" HeaderText="Section" />
        </columns>
        
        <EditRowStyle BackColor="#7C6F57" />
        <FooterStyle BackColor="#1C5E55" Font-Bold="True" ForeColor="White" />
        <HeaderStyle BackColor="#1C5E55" Font-Bold="True" ForeColor="White" />
        <PagerStyle BackColor="#666666" ForeColor="White" HorizontalAlign="Center" />
        <RowStyle BackColor="#E3EAEB" />
        <SelectedRowStyle BackColor="#C5BBAF" Font-Bold="True" ForeColor="#333333" />
        
    </asp:GridView>
    <br />
    <div>
    </div>
</asp:Content>
