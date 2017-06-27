<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Site1.Master" CodeBehind="UploadAgent.aspx.vb" Inherits="TYEACCOUNT.UploadAgent" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
    <script type="text/javascript" language="javascript">
       
    </script>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div class="download">

        <asp:ImageButton ImageUrl="~/images/excel-icon.png" ID="btnDownload" runat="server" Height="36px" Width="98px" ImageAlign="Middle" OnClick="btnDownload_Click" />
    </div>
    <table>
        <tr>
            <td>
                <asp:FileUpload ID="FileUpload1" runat="server" Width="468px" />
            </td>
        </tr>
        <%-- <tr>
            <td><asp:DropDownList ID="ddlPolicy" runat="server" Height="17px" Width="86px">
                <asp:ListItem Text="76" Value="76"></asp:ListItem>
                <asp:ListItem Text="77" Value="77"></asp:ListItem>
                <asp:ListItem Text="78" Value="78"></asp:ListItem>
                <asp:ListItem Text="79" Value="79"></asp:ListItem>
                <asp:ListItem Text="80" Value="80"></asp:ListItem>
                <asp:ListItem Text="81" Value="81"></asp:ListItem>
                <asp:ListItem Text="82" Value="82"></asp:ListItem>
                <asp:ListItem Text="83" Value="83"></asp:ListItem>
                </asp:DropDownList></td>
        </tr>--%>
        <tr>
            <td>
                <asp:DropDownList ID="ddltest" runat="server" Height="17px" Width="86px"></asp:DropDownList>
            </td>
        </tr>
        <tr>
            <td>
                <asp:Button ID="btnImportAgent" Text="Import Class" runat="server" Width="119px" />
                <br />
                <br />

            </td>
        </tr>
    </table>
    <asp:GridView ID="GridView2" runat="server">
    </asp:GridView>
    <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="false" BorderColor="#cccccc"
        OnRowDataBound="GridView1_RowDataBound" OnRowDeleting="GridView1_RowDeleting">
        <RowStyle BackColor="#E7E7FF" ForeColor="#4A3C8C" Font-Names="Tahoma" Font-Size="12px" />
        <PagerStyle BackColor="#E7E7FF" ForeColor="#4A3C8C" HorizontalAlign="Right" Font-Names="Tahoma" Font-Size="12px"/>
        <SelectedRowStyle BackColor="#738A9C" Font-Bold="True" ForeColor="#F7F7F7" Font-Names="Tahoma" Font-Size="12px"/>
        <HeaderStyle BackColor="#4A3C8C" Font-Bold="True" ForeColor="#F7F7F7" Font-Names="Tahoma" Font-Size="14px"/>
        <AlternatingRowStyle BackColor="#F7F7F7" />
        <Columns>
            <asp:TemplateField HeaderText="LotNo" HeaderStyle-Height="30px">
                <ItemTemplate>
                    <asp:Label ID="lblLotno" runat="server" Text='<%# DataBinder.Eval(Container, "DataItem.AGLOT") %>' CssClass="Padding">
                    </asp:Label>
                </ItemTemplate>
            </asp:TemplateField>
            <asp:TemplateField HeaderText="Agent">
                <ItemTemplate>
                    <asp:Label ID="lblAgent" runat="server" Text='<%# DataBinder.Eval(Container, "DataItem.AGNAME") %>' CssClass="Padding">
                    </asp:Label>
                </ItemTemplate>
            </asp:TemplateField>
            <asp:TemplateField HeaderText="Month/Year">
                <ItemTemplate>
                    <asp:Label ID="lblMinMax" runat="server" Text='<%# DataBinder.Eval(Container, "DataItem.MinMax")%>' CssClass="Padding">
                    </asp:Label>
                </ItemTemplate>
            </asp:TemplateField>
            <asp:TemplateField HeaderText="Add Date">
                <ItemTemplate>
                    <asp:Label ID="lblMY" runat="server" Text='<%# DataBinder.Eval(Container, "DataItem.AGADDT") %>' CssClass="Padding">
                    </asp:Label>
                </ItemTemplate>
        
            </asp:TemplateField>
            <asp:TemplateField HeaderText="Delete Lot" HeaderStyle-CssClass="PaddingH">
                <ItemTemplate>
                    <asp:LinkButton ID="btnDelete" runat="server" CommandName="Delete" Text="Delete" 
                        CommandArgument='<%# DataBinder.Eval(Container, "DataItem.AGLOT") %>'>
                     <div id ="divbtn" class="Deletebtn"></div>
                    </asp:LinkButton>
                </ItemTemplate>
                <ItemStyle HorizontalAlign="Center" />
            </asp:TemplateField>
        </Columns>
    </asp:GridView>
    <style>
         .PaddingH {
            padding:10px;
        }
        .Padding {
            padding:20px;
        }
        .Deletebtn {
            width: 40px;
            height:25px;
      
            background-image:url("../images/Del1.png");
            background-size:25px 25px;
            background-repeat:no-repeat;
            background-position:center center;
        }
        .Deletebtn:link {
            background-image:url("../images/Del1.png");
             background-size:20px 20px;
        }

        /* visited link */
        .Deletebtn:visited {
            background-color: green;
             background-size:25px 25px;
        }

        /* mouse over link */
        .Deletebtn:hover {
            /*background-color: hotpink;*/
             background-image:url("../images/Del2.png");
              background-size:25px 25px;
        }

        /* selected link */
        .Deletebtn:active {
            /*background-color: blue;*/
             background-image:url("../images/Del3.png");
              background-size:25px 25px;
        }
    </style>
</asp:Content>
