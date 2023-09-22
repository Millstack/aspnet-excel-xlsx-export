<%@ Page Language="C#" AutoEventWireup="true" CodeFile="ExcelXlsx.aspx.cs" Inherits="ExcelXlsx" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
        <div>
            <asp:Button ID="btnExp" runat="server" Text="Export" OnClick="btnExp_Click" />
        </div>
        <div class="table-responsive">
            <!--begin: Datatable -->
            <asp:GridView ShowHeaderWhenEmpty="true" ID="GridUser" runat="server" AutoGenerateColumns="false">
                <HeaderStyle CssClass="kt-shape-bg-color-3" />
                <Columns>
                    <asp:TemplateField ControlStyle-CssClass="col-xs-1" HeaderText="Sr.No">
                        <ItemTemplate>
                            <asp:HiddenField ID="id" runat="server" Value="id" />
                            <span>
                                <%#Container.DataItemIndex + 1%>
                            </span>
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:BoundField DataField="Name" HeaderText="Name" SortExpression="Name" />
                    <asp:BoundField DataField="Email" HeaderText="Email" SortExpression="Email" />
                    <asp:BoundField DataField="MobileNo" HeaderText="Mobile No" SortExpression="MobileNo" />
                    <asp:BoundField DataField="Password" HeaderText="Password" SortExpression="Password" />
                </Columns>
            </asp:GridView>
            <!--end: Datatable -->
        </div>
        
    </form>
</body>
</html>
