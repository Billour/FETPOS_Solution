<%@ Page Title="Home Page" Language="C#" MasterPageFile="~/Site.master" AutoEventWireup="true"
    CodeBehind="Default.aspx.cs" Inherits="ToolsWeb._Default" %>

<asp:Content ID="HeaderContent" runat="server" ContentPlaceHolderID="HeadContent">
</asp:Content>
<asp:Content ID="BodyContent" runat="server" ContentPlaceHolderID="MainContent">
    <table ><tr><td>Create Oracle Docx</td><td><asp:Button ID="btnCreate" 
            runat="server" Text="建立文件" onclick="btnCreate_Click"/></td></tr></table>
</asp:Content>
