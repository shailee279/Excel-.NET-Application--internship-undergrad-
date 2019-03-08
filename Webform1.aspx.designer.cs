<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="WebForm1.aspx.cs" Inherits="ExcelGeneration.WebForm1" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <meta charset="utf-8"/>
  <meta name="viewport" content="width=device-width, initial-scale=1"/>
     <link href="Styles/bootstrap-theme.css" rel="stylesheet"/>
      <link href="Styles/bootstrap-theme.min.css" rel="stylesheet"/>
       <link href="Styles/bootstrap.css" rel="stylesheet"/>
        <link href="Styles/bootstrap.min.css" rel="stylesheet"/>
       
   <script  type ="text/javascript" src="Scripts/bootstrap.js"></script>
   <script  type="text/javascript" src="Scripts/bootstrap.min.js"></script>

 <style>
body 
{  background-color:#cccccc;
   
 <%--  
  <%--<%-- width: 800px;
    
  <background-color: #99ccff ; padding: 200px;
    <%--border: 25px solid navy;
    margin: 25px;--%>--%>--%>--%>--%>
    
}



</style>
</head>
<body >
    <form id="form1" runat="server">
    <div class="form-group">
    <center>    
    <asp:Label runat="server" ID="label2" Text="LOGIN  PAGE" Font-Size="X-Large" 
            Font-Bold="True" Font-Underline="True" ForeColor="Black"></asp:Label><br /><br />
          
        <asp:Label ID="Label1" runat="server" Text="Username" Font-Bold="True" Font-Size="Large"></asp:Label><br />
        <asp:TextBox ID="TUsername" runat="server"  class="form-control" ></asp:TextBox><br />
        <asp:RequiredFieldValidator ID="RequiredFieldValidator1" runat="server" ErrorMessage="Username is required"  ForeColor="Red" ControlToValidate="TUsername"></asp:RequiredFieldValidator><br/>
        <asp:Label ID="Label3" runat="server" Text="Password" Font-Bold="True" Font-Size="Large"></asp:Label>
        <br />
        <asp:TextBox ID="TPassword" runat="server" TextMode="Password" class="form-control" BackColor="White"></asp:TextBox>
        <br />
        <asp:RequiredFieldValidator ID="RequiredFieldValidator2" runat="server" ErrorMessage="Password is required" ForeColor="Red" ControlToValidate="TPassword" ></asp:RequiredFieldValidator>
       <br />
        <asp:Label ID="Label4" runat="server" Text="Email ID" Font-Bold="True" Font-Size="Large"></asp:Label><br />
        <asp:TextBox ID="TEmail" runat="server"   class="form-control" TextMode="Email"></asp:TextBox><br />
        <asp:RequiredFieldValidator ID="RequiredFieldValidator3" runat="server" ErrorMessage="Email ID is required" ForeColor="Red"  ControlToValidate="TEmail"></asp:RequiredFieldValidator>
        <br /><br />
        
        <asp:Button ID="Button1" runat="server" Text="Submit" OnClick="btnSubmit_Click"   
          class=" btn btn-info btn-block"  />
        


    </center>
  

    </div>
    </form>
</body>
</html>