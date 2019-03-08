
<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="WebForm2.aspx.cs" Inherits="ExcelGeneration.WebForm2" %>

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
</head>
<style>
     body
     {background-color:white;
     }
     </style>


<body >
    <form id="form1" runat="server">
    <div class="form-group ">
    <center>
         <asp:ScriptManager ID="ScriptManager1" runat="server">
        </asp:ScriptManager>
        <asp:Label ID="Label1" runat="server" Text="Select File" ForeColor ="Black" Font-Bold="True" Font-Size="Large" ></asp:Label><br />
        <asp:FileUpload ID="FileUpload1" runat="server" AllowMultiple="true" 
         BackColor="White" ForeColor="Black" class=" form form-control"/><br /><br />
        <asp:RequiredFieldValidator ID="RequiredFieldValidator3" runat="server"  ForeColor="Red" ErrorMessage="Please select file" ControlToValidate="FileUpload1" ValidationGroup="ab"></asp:RequiredFieldValidator><br /><br />
   <asp:Label ID="LblExcel" runat="server" Text="Enter Excel File Name" ForeColor="Black" Font-Bold="True" Font-Size="Large"></asp:Label><br />
        <asp:TextBox ID="TextExcel" runat="server" class ="form-control"></asp:TextBox><br />
            <asp:RequiredFieldValidator ID="RequiredFieldValidator1" runat="server" 
            ErrorMessage="File Name required" ValidationGroup="ab" 
            ControlToValidate="TextExcel" ForeColor="Red"></asp:RequiredFieldValidator>
        <br /><br /><br />
        <asp:Label ID="Label3" runat="server" Text="Enter Email Address" ForeColor="Black" Font-Bold="True" Font-Size="Large"></asp:Label><br />
        <asp:TextBox ID="TextEmail" runat="server"  TextMode="Email" class="form-control"></asp:TextBox><br />
        <asp:RequiredFieldValidator ID="RequiredFieldValidator2" runat="server"  ValidationGroup=
        "ab" ErrorMessage="Email ID is required" ControlToValidate="TextEmail" ForeColor="Red"></asp:RequiredFieldValidator><br />
       <br />
        <asp:Button ID="Button3" runat="server" Text="Start"  OnClick="Start_Click" ValidationGroup="ab"  class=" btn btn-info btn-block"/>&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<br />
        <asp:Button ID="Button4" runat="server" Text="Cancel"   OnClick="Cancel_Click" class=" btn btn-info btn-block"/><br /><br />
       <%-- <asp:Label ID="Label5" runat="server" Text="Excel file saved in Folder" Visible="false"></asp:Label><br /><br />--%>


       
        <%--<asp:GridView ID="gvDetails" CellPadding="5" runat="server" 
            AutoGenerateColumns="False">
<Columns>
<asp:BoundField DataField="Text" HeaderText="FileName" />
</Columns>
<HeaderStyle BackColor="Black" Font-Bold="true" ForeColor="White" />
</asp:GridView>--%>
        <asp:LinkButton ID="LinkButton1" runat="server" Text="Log Out" OnClick="Log_Click"  class="btn-link btn-block" ></asp:LinkButton><br /><br />
       


        <br />
        <br />
      

        </center>
    </div>
    </form>   </body>
</html>