<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="inicio.aspx.vb" Inherits="GripViewaExcel.inicio" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title></title>
</head>
<body>

    <form id="form1" runat="server">
        <asp:ScriptManager ID="ScriptManager1" runat="server">
        </asp:ScriptManager>
        <asp:UpdatePanel ID="UpdatePanel1" runat="server">

            <Triggers>
                <asp:PostBackTrigger ControlID="btnExportar" />
                
            </Triggers>

            <ContentTemplate>

                <div>
                    Prueba de Exportar Datos de GripView a Excel<br />

                    <br />
                    base <a data-linktype="external" href="https://github.com/Microsoft/sql-server-samples/releases/download/adventureworks/AdventureWorks2019.bak" style="box-sizing: inherit; background-color: rgb(255, 255, 255); outline-color: inherit; color: var(--theme-visited); cursor: pointer; text-decoration: underline; overflow-wrap: break-word; outline-style: initial; outline-width: 0px; font-family: &quot;Segoe UI&quot;, SegoeUI, &quot;Helvetica Neue&quot;, Helvetica, Arial, sans-serif; font-size: 14px; font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-align: left; text-indent: 0px; text-transform: none; white-space: normal; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px;">AdventureWorks2019.bak</a><br /> <a href="https://docs.microsoft.com/en-us/sql/samples/adventureworks-install-configure?view=sql-server-ver15&amp;tabs=ssms">
                    https://docs.microsoft.com/en-us/sql/samples/adventureworks-install-configure?view=sql-server-ver15&amp;tabs=ssms</a><br />

                    <br />
                    <br />
                 
            <asp:Button ID="btnExportar" OnClick="btnExportar_Click" runat="server" Text="Exportar Arreglado" />
                    <br />
                </div>
                <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" DataKeyNames="BusinessEntityID" DataSourceID="SqlDataSource1">
                    <Columns>
                        <asp:BoundField DataField="PersonType" HeaderText="PersonType" SortExpression="PersonType" />
                        <asp:BoundField DataField="Title" HeaderText="Title" SortExpression="Title" HtmlEncode="False" />
                        <asp:BoundField DataField="FirstName" HeaderText="FirstName" SortExpression="FirstName" />
                        <asp:BoundField DataField="MiddleName" HeaderText="MiddleName" SortExpression="MiddleName" />
                        <asp:BoundField DataField="LastName" HeaderText="LastName" SortExpression="LastName1" AccessibleHeaderText="Apellido" />
                    </Columns>
                </asp:GridView>
                <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:AdventureWorks2019ConnectionString %>" SelectCommand="/****** Script para el comando SelectTopNRows de SSMS  ******/
                                    SELECT TOP (1000) [BusinessEntityID]
                                          ,[PersonType]
                                          ,[Title]
                                          ,[FirstName]
                                          ,[MiddleName]
                                          ,[LastName]
                                          ,[Suffix]
                                          ,[EmailPromotion]
                                     FROM [AdventureWorks2019].[Person].[Person]"></asp:SqlDataSource>

            </ContentTemplate>


        </asp:UpdatePanel>

    </form>
</body>
</html>
