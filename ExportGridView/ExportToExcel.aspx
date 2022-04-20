<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="ExportToExcel.aspx.cs" Inherits="ExportGridView.ExportToExcel" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Export GridView To Excel</title>
</head>
<body>
    <link href="Styles/StyleSheet.css" rel="stylesheet" />
    <form id="form1" runat="server">
        <div>
            <table align="center">
                <tr>
                    <td class="tableHeader">Employee Master Data
                    </td>
                </tr>
                <tr>
                    <td>
                        <asp:GridView ID="gvEmployeeData" runat="server" AutoGenerateColumns="False" BackColor="White" BorderColor="#CC9966" BorderStyle="None" BorderWidth="1px" CellPadding="4">
                            <Columns>
                                <asp:BoundField HeaderText="#" DataField="Rownumber" ItemStyle-HorizontalAlign="Center">
                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                </asp:BoundField>
                                <asp:BoundField HeaderText="NAME" DataField="NAME" ItemStyle-HorizontalAlign="Left">
                                    <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                </asp:BoundField>
                                <asp:BoundField HeaderText="DESIGNATION" DataField="DESIGNATION" ItemStyle-HorizontalAlign="Left">
                                    <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                </asp:BoundField>
                                <asp:BoundField HeaderText="SALARY" DataField="SALARY" ItemStyle-HorizontalAlign="Right">
                                    <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                </asp:BoundField>
                                <asp:BoundField HeaderText="STATE" DataField="STATE" ItemStyle-HorizontalAlign="Left">
                                    <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                </asp:BoundField>
                                <asp:BoundField HeaderText="COUNTRY" DataField="COUNTRY" ItemStyle-HorizontalAlign="Left">
                                    <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                </asp:BoundField>
                            </Columns>
                            <FooterStyle BackColor="#FFFFCC" ForeColor="#330099" />
                            <HeaderStyle BackColor="#990000" Font-Bold="True" ForeColor="#FFFFCC" />
                            <PagerStyle BackColor="#FFFFCC" ForeColor="#330099" HorizontalAlign="Center" />
                            <RowStyle BackColor="White" ForeColor="#330099" />
                            <SelectedRowStyle BackColor="#FFCC66" Font-Bold="True" ForeColor="#663399" />
                            <SortedAscendingCellStyle BackColor="#FEFCEB" />
                            <SortedAscendingHeaderStyle BackColor="#AF0101" />
                            <SortedDescendingCellStyle BackColor="#F6F0C0" />
                            <SortedDescendingHeaderStyle BackColor="#7E0000" />
                        </asp:GridView>
                    </td>
                </tr>
                <tr>
                    <td>
                        <asp:Label ID="lblMessage" runat="server" Font-Bold="true" ForeColor="Green"></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td>
                        <hr />
                    </td>
                </tr>
                <tr>
                    <td style="text-align: center">
                        <asp:Button ID="btnExport" runat="server" Text="Export to Excel" CssClass="button" OnClick="btnExport_Click" />
                    </td>
                </tr>
            </table>
        </div>
    </form>
</body>
</html>
