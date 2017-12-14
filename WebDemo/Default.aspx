<%@ Page Title="Home Page" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="WebDemo._Default" %>

<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">

    <table border="1" id="tblDatabases"  style="width: auto">
        <asp:Repeater ID="DatabaseList" runat="server" Visible="True">
            <HeaderTemplate>
                <thead>
                    <tr>
                        <th>
                            Database Name
                        </th>
                    </tr>
                </thead>
            </HeaderTemplate>
            <ItemTemplate>
                <tr>
                    <td>
                        <span title='<%# DataBinder.Eval(Container.DataItem, "state_desc") %>'>
                            <%# DataBinder.Eval(Container.DataItem, "name") %>
                        </span>
                    </td>
                </tr>
            </ItemTemplate>
        </asp:Repeater>
    </table>



</asp:Content>
