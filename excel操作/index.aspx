<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="index.aspx.cs" Inherits="excel操作.index" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title>Excel操作</title>
</head>
<body>
    <form id="form1" runat="server">
        <div>
            <asp:RadioButtonList ID="RadioButtonList1" runat="server"></asp:RadioButtonList>
            <%--第一步  上传文件--%>
            <asp:FileUpload ID="FileUpload1" runat="server" />
            <asp:Button Text="上传" ID="BtnUp" OnClick="BtnUp_Click" runat="server" />
            <br />
            <asp:Image ID="Image1" runat="server" />
            <asp:GridView ID="GridView1" runat="server"></asp:GridView>



            <asp:Button Text="下载" ID="BtnDown" OnClick="BtnDown_Click" runat="server" />
             <asp:Button Text="预览" ID="Button1" OnClick="Button1_Click" runat="server" />
             <asp:GridView ID="GridView2" runat="server"></asp:GridView>

        </div>
    </form>
</body>
</html>
