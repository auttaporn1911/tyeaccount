<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Site1.Master" CodeBehind="uploadsales.aspx.vb" Inherits="TYEACCOUNT.WebForm2" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <h2>Upload Sales Data Master</h2>
    <div class="download">

        <asp:ImageButton ImageUrl="~/images/excel-icon.png" ID="btnDownload" runat="server" Height="36px" Width="90px" ImageAlign="Middle" OnClick="btnDownload_Click" />
    </div>
    <table>
        <tr>
            <td style="width:170px">Select Files :</td>
            <td>
                <asp:FileUpload ID="FileUpload1" runat="server" Width="498px" />
                <br />
                <asp:Label ID="lbFileName" runat="server"></asp:Label>
                <br />
            </td>
        </tr>
        <tr>
            <td style="width:170px">Select Sheet :</td>
            <td>
                <asp:DropDownList ID="ddlSheet" runat="server" Width="148px"></asp:DropDownList>
                <asp:Button ID="btnLoad" Text="Refresh" runat="server" Width="68px" />
            </td>
        </tr>

        <tr>
            <td>
                <asp:Button ID="btnImport" Text="Import Sales" runat="server" OnClientClick="DisplayLoadingModal();" />

                <br />
                <br />
                

            </td>
        </tr>
    </table>
     <fieldset class="fsStandard" style="padding-bottom: 6px; width: 97%;">
    <legend>Upload Detail</legend>
         
    <asp:GridView ID="gvLot" runat="server" AutoGenerateColumns="False" CellPadding="4" EnableModelValidation="True" ForeColor="#333333" GridLines="None">
                    <AlternatingRowStyle BackColor="White" />
                    <Columns>
                        <asp:BoundField DataField="STARTDATE" HeaderText="Start Date" />
                        <asp:BoundField DataField="ENDDATE" HeaderText="End Date" />
                        
                        <asp:BoundField DataField="QTY" HeaderText="Number of Record" DataFormatString="{0:#,###}" />
                        <asp:BoundField DataField="Amount" HeaderText="Amount"  DataFormatString="{0:#,###.00}"/>
                        <asp:BoundField DataField="LOTNO" HeaderText="Lot Number" />
                        <asp:TemplateField ItemStyle-Width="35px" ItemStyle-HorizontalAlign="Center">
                            <ItemTemplate>
                                <asp:LinkButton ID="imgEdit" OnClick="delete_click" runat="server"><asp:Image ImageUrl="images/erase.png" runat="server"  />
                                </asp:LinkButton>
                            </ItemTemplate>

                            <ItemStyle HorizontalAlign="Center" Width="35px"></ItemStyle>
                        </asp:TemplateField>

                    </Columns>
                    <EditRowStyle BackColor="#2461BF" />
                    <FooterStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
                    <HeaderStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
                    <PagerStyle BackColor="#2461BF" ForeColor="White" HorizontalAlign="Center" />
                    <RowStyle BackColor="#EFF3FB" />
                    <SelectedRowStyle BackColor="#D1DDF1" Font-Bold="True" ForeColor="#333333" />
                </asp:GridView></fieldset>
                <br />
                <asp:GridView ID="gvErr" runat="server" BackColor="White" BorderColor="#CC9966" BorderStyle="None" BorderWidth="1px" CellPadding="4" EnableModelValidation="True">
                    <FooterStyle BackColor="#FFFFCC" ForeColor="#330099" />
                    <HeaderStyle BackColor="#990000" Font-Bold="True" ForeColor="#FFFFCC" />
                    <PagerStyle BackColor="#FFFFCC" ForeColor="#330099" HorizontalAlign="Center" />
                    <RowStyle BackColor="White" ForeColor="#330099" />
                    <SelectedRowStyle BackColor="#FFCC66" Font-Bold="True" ForeColor="#663399" />
                </asp:GridView>
                <br />
                <br />
</asp:Content>
