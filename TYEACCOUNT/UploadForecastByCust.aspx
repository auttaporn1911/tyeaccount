<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Site1.Master" CodeBehind="UploadForecastByCust.aspx.vb" Inherits="TYEACCOUNT.WebForm9" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <h2>Upload Forecast By Plan</h2>
    <div class="download">

        <asp:ImageButton ImageUrl="~/images/excel-icon.png" ID="btnDownload" runat="server" Height="36px" Width="98px" ImageAlign="Middle" OnClick="btnDownload_Click" />
    </div>
    <table>
        <tr>
            <td>Files Upload :</td>
            <td>
                <asp:FileUpload ID="FileUpload1" runat="server" Width="468px" />
            </td>
        </tr>
        <tr>
            <td>Policy Year : </td>
            <td>
                <asp:DropDownList ID="ddlPolicy" runat="server" Height="17px" Width="86px">
                    <asp:ListItem Text="74" Value="74"></asp:ListItem>
                    <asp:ListItem Text="75" Value="75"></asp:ListItem>
                    <asp:ListItem Text="76" Value="76"></asp:ListItem>
                    <asp:ListItem Text="77" Value="77"></asp:ListItem>
                    <asp:ListItem Text="78" Value="78"></asp:ListItem>
                    <asp:ListItem Text="79" Value="79"></asp:ListItem>
                    <asp:ListItem Text="80" Value="80"></asp:ListItem>
                    <asp:ListItem Text="81" Value="81"></asp:ListItem>
                    <asp:ListItem Text="82" Value="82"></asp:ListItem>
                    <asp:ListItem Text="83" Value="83"></asp:ListItem>
                </asp:DropDownList></td>
        </tr>
        <tr>
            <td></td>
            <td>
                <asp:Button ID="btnImportForecast" Text="Import Class" runat="server" Width="119px" />


            </td>

        </tr>
    </table>
    <br />
    <fieldset class="fsStandard" style="padding-bottom: 6px; width: 60%;">
        <legend>Upload Detail</legend>
        <asp:GridView ID="gvLot" runat="server" AutoGenerateColumns="False" CellPadding="3" EnableModelValidation="True" BackColor="White" BorderColor="#CCCCCC" BorderStyle="None" BorderWidth="1px">
            <Columns>
                <asp:BoundField DataField="STARTDATE" HeaderText="Start Date" />
                <asp:BoundField DataField="ENDDATE" HeaderText="End Date" />

                <asp:BoundField DataField="LOTNO" HeaderText="Lot Number" />
                <asp:TemplateField ItemStyle-Width="35px" ItemStyle-HorizontalAlign="Center">
                    <ItemTemplate>
                        <asp:LinkButton ID="imgEdit" OnClick="delete_click" runat="server"><asp:Image ImageUrl="images/erase.png" runat="server"  />
                        </asp:LinkButton>
                    </ItemTemplate>

                    <ItemStyle HorizontalAlign="Center" Width="35px"></ItemStyle>
                </asp:TemplateField>

            </Columns>
            <FooterStyle BackColor="White" ForeColor="#000066" />
            <HeaderStyle BackColor="#006699" Font-Bold="True" ForeColor="White" />
            <PagerStyle BackColor="White" ForeColor="#000066" HorizontalAlign="Left" />
            <RowStyle ForeColor="#000066" />
            <SelectedRowStyle BackColor="#669999" Font-Bold="True" ForeColor="White" />
        </asp:GridView>
        <br />
    </fieldset>
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
    <div id="note1" class="note">
        1.sheet name must be "Plan Sale by Customer"
    </div>
    <br />
    <br />
    <br />
    <br />
</asp:Content>
