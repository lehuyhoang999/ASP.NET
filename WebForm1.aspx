<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="WebForm1.aspx.cs" Inherits="WebApplicationDemo.WebForm1" %>

<%@ Register Assembly="DevExpress.Web.ASPxSpreadsheet.v19.2, Version=19.2.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web.ASPxSpreadsheet" TagPrefix="dx" %>

<%@ Register Assembly="DevExpress.Web.v19.2, Version=19.2.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web" TagPrefix="dx" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
        <div>
           
            
            <asp:DropDownList ID="DropDownList1" runat="server"></asp:DropDownList>
            <dx:ASPxButton ID="ASPxButton1" runat="server" Text="ASPxButton" OnClick="btnClick"></dx:ASPxButton>
            <dx:ASPxSpreadsheet ID="Spreadsheet" ClientInstanceName="spreadsheet" runat="server" Width="100%" Height="700px" Visible="false" ActiveTabIndex="0" ShowConfirmOnLosingChanges="false"></dx:ASPxSpreadsheet>
            <dx:ASPxGridView ID="grid" runat="server" Width="100%" AutoGenerateColumns="False">
                <Toolbars>
                    <dx:GridViewToolbar>
                        <SettingsAdaptivity Enabled="true" EnableCollapseRootItemsToIcons="true" />
                        <Items>
                            <dx:GridViewToolbarItem Command="ExportToPdf" />
                            <dx:GridViewToolbarItem Command="ExportToXls" />
                            <dx:GridViewToolbarItem Command="ExportToXlsx" />
                            <dx:GridViewToolbarItem Command="ExportToDocx" />
                            <dx:GridViewToolbarItem Command="ExportToRtf" />
                            <dx:GridViewToolbarItem Command="ExportToCsv" />
                        </Items>
                    </dx:GridViewToolbar>
                </Toolbars>
                <Columns>
                <dx:GridViewDataColumn Caption="Product Name" FieldName="ProductName"/>
                <dx:GridViewDataColumn Caption="Company Name" FieldName="CompanyName" />
                <dx:GridViewDataColumn Caption="Order Date" FieldName="OrderDate" />
                <%--<dx:GridViewDataTextColumn Caption="Product Amount" FieldName="ProductAmount" ReadOnly="True">
                    <PropertiesTextEdit DisplayFormatString="c" />
                </dx:GridViewDataTextColumn>--%>
                </Columns>
                <Settings ShowGroupPanel="True" ShowFooter="True" ShowFilterRow="True"/>
                <SettingsExport EnableClientSideExportAPI="true" ExcelExportMode="WYSIWYG" />
                <%--<GroupSummary>
                    <dx:ASPxSummaryItem FieldName="ProductAmount" SummaryType="Sum" />
                    <dx:ASPxSummaryItem FieldName="CompanyName" SummaryType="Count" />
                </GroupSummary>
                <TotalSummary>
                    <dx:ASPxSummaryItem FieldName="CompanyName" SummaryType="Count" />
                    <dx:ASPxSummaryItem FieldName="ProductAmount" SummaryType="Sum" />
                </TotalSummary>--%>
            </dx:ASPxGridView>
        </div>
        
    </form>
</body>
</html>
