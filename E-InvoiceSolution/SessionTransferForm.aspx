<%@ Page Language="C#" AutoEventWireup="true" CodeFile="SessionTransferForm.aspx.cs"
    Inherits="_SessionTransferForm" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <title>SessionTransfer</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="C#">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
</head>
<body>
    <form id="frm1" name="frm1" method="post" action="<%=Request.QueryString["url"] %>" >
        <%
            foreach (string key in Request.Form)
            {
        %>
        <input type="hidden" id="<%=key%>" name="<%=key%>" value="<%=Request.Form[key].ToString()%>" />
        <%
            }
        %>
    </form>
<script language="JavaScript">
document.frm1.submit();
</script>
</body>
</html>
