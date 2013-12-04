<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="UserImport.aspx.cs" Inherits="SitecoreExtension.SitecoreUserImporter.UserImport" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
        <h2>Developer tool / only to be used by trained user</h2>
		<div>
		    <asp:textbox runat="server" id="txtUserDomain" Text="sitecore"/>
			<asp:checkbox runat="server" id="chkReplaceEmails" text="Replaces original emails with testx@1508.dk" />
			<br />
			<asp:label runat="server" id="lblMaxImports" text="Max antal brugere importeret: " />
			<asp:textbox runat="server" id="txtMaxImports" />
			<br />
			<asp:FileUpload id="fileImportFile" runat="server" />
			<asp:button id="btnUpload" runat="server" text="Upload" onclick="btnUpload_Click" /> (<%= UploadCount %> user uploaded)
			<br />
			<asp:button id="btnBuildUserList" runat="server" text="Build UserList" onclick="btnBuildUserList_Click" />
			<br />
			<asp:button id="btnValidateRoles" runat="server" text="Validate Roles" onclick="btnValidateRoles_Click"/>
			<br />
			<asp:button id="btnTestUserImport" runat="server" text="Test User Import" onclick="btnTestUserImport_Click" />
			<br />
			<asp:button id="btnShowExistingUsers" runat="server" text="Show Existing Users" onclick="btnShowExistingUsers_Click" style="height: 26px" />
			<br />
			<asp:button id="btnImportUsers" runat="server" text="Import Users" onclick="btnImportUsers_Click" />
			<br />
			<asp:button id="btnExportUsersInList" runat="server" text="Build Sitecore Package (only with users in the list)" onclick="btnExportUsersInList_Click" />
			<br />
			<asp:button id="btnDeleteUsersInList" runat="server" text="Delete users in list" onclick="btnDeleteUsersInList_Click" />
			<br />
			<label>Notification-Mail Template</label>
			<br />
			<asp:textbox id="txtNotificationMailTemplate" runat="server" height="300px" width="80%" />
			<asp:button id="btnSendNotifications" runat="server" text="Send Notifications" onclick="btnSendNotifications_Click" />
			<br />
		</div>
		<div id="divDebug" runat="server">
		</div>
    </form>
</body>
</html>
