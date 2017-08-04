<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Default.aspx.cs" Inherits="mtlList._Default" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <style type="text/css">
        .style1
        {
            width: 92px;
        }
        .style2
        {
            font-size: xx-large;
        }
        .style3
        {
            width: 204px;
        }
        .style4
        {
            font-size: small;
        }
        .style5
        {
            width: 92px;
            height: 29px;
        }
        .style6
        {
            width: 204px;
            height: 29px;
        }
        .style7
        {
            height: 29px;
        }
        .style8
        {
            width: 92px;
            height: 13px;
        }
        .style9
        {
            width: 204px;
            height: 13px;
        }
        .style10
        {
            height: 13px;
        }
        .auto-style1 {
            width: 81px;
        }
        .auto-style2 {
            height: 29px;
            width: 81px;
        }
        .auto-style3 {
            height: 13px;
            width: 81px;
        }
        .auto-style4 {
            width: 151px;
        }
        .auto-style5 {
            width: 151px;
            height: 29px;
        }
        .auto-style6 {
            width: 151px;
            height: 13px;
        }
        .auto-style7 {
            width: 160px;
        }
        .auto-style8 {
            height: 29px;
            width: 160px;
        }
        .auto-style9 {
            height: 13px;
            width: 160px;
        }
    </style>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <strong><span class="style2">對帳用進料狀況表</span></strong><span class="style2"><strong><br />
        </strong></span>
        <table style="width: 100%;">
            <tr>
                <td bgcolor="#FFFF99" class="auto-style4">
                    起始日期：(必填)</td>
                <td class="style3">
                    <asp:TextBox ID="txtDate1" runat="server" TabIndex="5">20170703</asp:TextBox>
                    &nbsp;<br />
                    <span class="style4"><strong>格式 20131205</strong></span></td>
                <td bgcolor="#3399FF" class="auto-style1">
                    &nbsp;
                    異動類型</td>
                <td class="auto-style7">
                    <asp:DropDownList ID="ddlMvt" runat="server" TabIndex="30">
                        <asp:ListItem>請選擇</asp:ListItem>
                        <asp:ListItem>103</asp:ListItem>
                        <asp:ListItem>104</asp:ListItem>
                        <asp:ListItem>105</asp:ListItem>
                    </asp:DropDownList>
                </td>
                <td>
                    &nbsp;</td>
            </tr>
            <tr>
                <td bgcolor="#FFFF99" class="auto-style4">
                    結束日期：(必填)</td>
                <td class="style3">
                    <asp:TextBox ID="txtDate2" runat="server" TabIndex="6">20170710</asp:TextBox>
                    <br />
                    <span class="style4"><strong>格式 20131205</strong></span></td>
                <td class="auto-style1">
                    &nbsp;
                    </td>
                <td class="auto-style7">
                    &nbsp;</td>
                <td>
                    &nbsp;</td>
            </tr>
            <tr>
                <td class="auto-style5">
                    &nbsp;</td>
                <td class="style6">
                    &nbsp;</td>
                <td class="auto-style2" bgcolor="#3399FF">
                    廠商名稱 %</td>
                <td class="auto-style8">
                    <asp:TextBox ID="txtVendorName" runat="server" TabIndex="41"></asp:TextBox>
                </td>
                <td class="style7">
                    &nbsp;</td>
            </tr>
            <tr>
                <td bgcolor="#3399FF" class="auto-style5">
                    採購單號</td>
                <td class="style6">
                    <asp:TextBox ID="txtPO" runat="server" TabIndex="10"></asp:TextBox>
                </td>
                <td class="auto-style2">
                    &nbsp;</td>
                <td class="auto-style8">
                    &nbsp;</td>
                <td class="style7">
                    &nbsp;</td>
            </tr>
            <tr>
                <td bgcolor="#3399FF" class="auto-style6">
                    物料號碼</td>
                <td class="style9">
                    <asp:TextBox ID="txtMaterial" runat="server" TabIndex="11"></asp:TextBox>
                </td>
                <td class="auto-style3" bgcolor="#3399FF">
                    物料文件</td>
                <td class="auto-style9">
                    <asp:TextBox ID="txtMtlDoc" runat="server" TabIndex="42"></asp:TextBox>
                </td>
                <td class="style10">
                    &nbsp;</td>
            </tr>
            <tr>
                <td class="auto-style6">
                    &nbsp;</td>
                <td class="style9">
                    &nbsp;</td>
                <td class="auto-style3" bgcolor="#3399FF">
                    參考文件</td>
                <td class="auto-style9">
                    <asp:TextBox ID="txtRefDoc" runat="server" TabIndex="43"></asp:TextBox>
                </td>
                <td class="style10">
                    &nbsp;</td>
            </tr>
            <tr>
                <td class="auto-style6" bgcolor="#3399FF">
                    待驗</td>
                <td class="style9">
                    <asp:CheckBox ID="cbZflag" runat="server" TabIndex="20" Text="待驗資料" />
                </td>
                <td class="auto-style3">
                    &nbsp;</td>
                <td class="auto-style9">
                    &nbsp;</td>
                <td class="style10">
                    &nbsp;</td>
            </tr>
            <tr>
                <td class="auto-style4">
                    &nbsp;</td>
                <td class="style3">
                    <asp:Button ID="btnQry" runat="server" OnClick="btnQry_Click" Text="查詢" Style="height: 26px" TabIndex="51" />
                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                    <asp:Button ID="btnClr" runat="server" OnClick="btnClr_Click" Text="重置" TabIndex="52" />
                &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                    </td>
                <td class="auto-style1">
                    <asp:Button ID="btnConvert" runat="server" onclick="btnConvert_Click" 
                        Text="轉EXCEL" Visible="False" TabIndex="53" />
                </td>
                <td class="auto-style7">
                    &nbsp;</td>
                <td>
                    &nbsp;</td>
            </tr>
        </table>
                    <asp:Label ID="lblMsg" runat="server"></asp:Label>
        <br />
        <asp:GridView ID="gvData" runat="server" BackColor="White" 
            BorderColor="#999999" BorderStyle="None" BorderWidth="1px" CellPadding="3" 
            GridLines="Vertical">
            <AlternatingRowStyle BackColor="#DCDCDC" />
            <FooterStyle BackColor="#CCCCCC" ForeColor="Black" />
            <HeaderStyle BackColor="#000084" Font-Bold="True" ForeColor="White" />
            <PagerStyle BackColor="#999999" ForeColor="Black" HorizontalAlign="Center" />
            <RowStyle BackColor="#EEEEEE" ForeColor="Black" />
            <SelectedRowStyle BackColor="#008A8C" Font-Bold="True" ForeColor="White" />
            <SortedAscendingCellStyle BackColor="#F1F1F1" />
            <SortedAscendingHeaderStyle BackColor="#0000A9" />
            <SortedDescendingCellStyle BackColor="#CAC9C9" />
            <SortedDescendingHeaderStyle BackColor="#000065" />
        </asp:GridView>
        <br />
        <br />
    <asp:HyperLink ID="hl" runat="server" NavigateUrl="/sap/">回SAP報表畫面</asp:HyperLink>
    </div>
    </form>
</body>
</html>
