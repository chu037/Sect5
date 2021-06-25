<%@ Page Language="VB" ContentType="text/html" ResponseEncoding="utf-8" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<%@ Register TagPrefix="MM" Namespace="DreamweaverCtrls" Assembly="DreamweaverCtrls,version=1.0.0.0,publicKeyToken=836f606ede05d46a,culture=neutral" %>

<MM:DataSet 
id="DataSet1"
runat="Server"
IsStoredProcedure="false"
ConnectionString='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_STRING_s01") %>'
DatabaseType='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_DATABASETYPE_s01") %>'
CommandText='<%# "SELECT * FROM s01_kc_calen WHERE m_date between ? and ? AND m_messeng LIKE ? AND m_note LIKE ? ORDER BY m_date DESC" %>'
Debug="true" PageSize="50"
>
  <Parameters>
    <Parameter  Name="@m_date"  Value='<%# IIf((Request.QueryString("m_date") <> Nothing), Request.QueryString("m_date"), "#1/1/2008#") %>'  Type="Date"   />
    <Parameter  Name="@m_date1"  Value='<%# IIf((Request.QueryString("m_date1") <> Nothing), Request.QueryString("m_date1"), "#12/31/2099#") %>'  Type="Date"   />

    <Parameter  Name="@m_messeng"  Value='<%# "%" + (IIf((Request.QueryString("m_messeng") <> Nothing), stra1(Request.QueryString("m_messeng")), "")) + "%" %>'  Type="WChar"   />
    <Parameter  Name="@m_note"  Value='<%# "%" + (IIf((Request.QueryString("m_note") <> Nothing), stra1(Request.QueryString("m_note")), "")) + "%" %>'  Type="WChar"   />

  </Parameters>
</MM:DataSet>
<MM:DataSet 
id="DataSet2"
runat="Server"
IsStoredProcedure="false"
ConnectionString='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_STRING_s01") %>'
DatabaseType='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_DATABASETYPE_s01") %>'
CommandText='<%# "SELECT * FROM s01_kc_calen WHERE m_date between ? and ? AND m_messeng LIKE ? AND m_note LIKE ? ORDER BY m_date DESC" %>'
Debug="true"
>
  <Parameters>
    <Parameter  Name="@m_date"  Value='<%# IIf((Request.QueryString("m_date") <> Nothing), Request.QueryString("m_date"), "#1/1/2008#") %>'  Type="Date"   />
    <Parameter  Name="@m_date1"  Value='<%# IIf((Request.QueryString("m_date1") <> Nothing), Request.QueryString("m_date1"), "#12/31/2099#") %>'  Type="Date"   />

    <Parameter  Name="@m_messeng"  Value='<%# "%" + (IIf((Request.QueryString("m_messeng") <> Nothing), stra1(Request.QueryString("m_messeng")), "")) + "%" %>'  Type="WChar"   />
    <Parameter  Name="@m_note"  Value='<%# "%" + (IIf((Request.QueryString("m_note") <> Nothing), stra1(Request.QueryString("m_note")), "")) + "%" %>'  Type="WChar"   />

  </Parameters>
</MM:DataSet>
<MM:PageBind runat="server" PostBackBind="true" />
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>行程查詢結果</title>
<script language="VB" runat="server">
Sub page_load(sender As Object, e As EventArgs)
end sub  
  
 function showdate(s01time)
 if s01time <> ""
 showdate = FormatDateTime(s01time, DateFormat.ShortDate)
 end if
 end function
 function showtime(t)
 if t <> ""
 showtime = FormatDateTime(t, DateFormat.Shorttime)
 end if
 end function
 function showmessage(vt2,vt3)
 if vt2 <> ""
 showmessage = vt2 & "<br/>" & vt3 & "小時"
 end if
 end function
 Protected Sub Button2_Click(ByVal sender As Object, ByVal e As System.EventArgs) 

        Dim ExportWriter As New IO.StringWriter()
        Dim ExportHtmlTextWriter As New HtmlTextWriter(ExportWriter)
        Dim xlsApp As Object  
        Dim strValue As String  
		'DataGrid1.AllowPaging = False
        'DataGrid1.AllowSorting = False
        DataGrid2.visible = true
		DisableControls(DataGrid2)
        DataGrid2.RenderControl(ExportHtmlTextWriter)

        Response.AppendHeader("Content-Disposition", "attachment;filename=" + Date.Today.ToString("yyyy-MM-dd") + ".xls")
 'xlsApp = CreateObject("Excel.Application")   
 'xlsApp.Workbooks.Open("C:\Book1.xls")   

        Response.ContentEncoding = System.Text.Encoding.GetEncoding("big5")
        Response.ContentType = "application/vnd.ms-excel"
        Response.Write(ExportWriter.ToString())
        DataGrid2.visible = False
		Response.End()
    End Sub

    Protected Sub DisableControls(ByVal control As Control)
        '處理GridView裡的控制項，將其變成literal
        Dim i As Integer = 0
        Do While (i < control.Controls.Count)
            Dim current As Control = control.Controls(i)
            If (TypeOf current Is LinkButton) Then
                control.Controls.Remove(current)
                control.Controls.AddAt(i, New LiteralControl(CType(current, LinkButton).Text))
            ElseIf (TypeOf current Is ImageButton) Then
                control.Controls.Remove(current)
                control.Controls.AddAt(i, New LiteralControl(CType(current, ImageButton).AlternateText))
            ElseIf (TypeOf current Is HyperLink) Then
                control.Controls.Remove(current)
                control.Controls.AddAt(i, New LiteralControl(CType(current, HyperLink).Text))
            ElseIf (TypeOf current Is DropDownList) Then
                control.Controls.Remove(current)
                control.Controls.AddAt(i, New LiteralControl(CType(current, DropDownList).SelectedItem.Text))
            ElseIf (TypeOf current Is CheckBox) Then
                control.Controls.Remove(current)
                control.Controls.AddAt(i, New LiteralControl(CType(current, CheckBox).Checked))
                'TODO: Warning!!!, inline IF is not supported ?
            End If

            If current.HasControls Then
                DisableControls(current)
            End If
            i = (i + 1)
        Loop
    End Sub
function stra1(t01)
 if t01 <> ""  
  dim i as integer
  dim str2 as string
  dim v
   v = split(t01,",")
   str2 = "" 
	for i = 0 to ubound(v)
	str2 &= chr(v(i))
	next
   stra1 = str2
  end if
end function   	
	
	


</script>
<script language = "JavaScript">
<!--

function w_back()
{
location.href = "kc_cal.aspx"
}
function MM_goToURL() { //v3.0
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"'");
}

//-->
</Script>
<style type="text/css">
<!--
.style2 {color: #0066FF}
.style3 {color: #CC3300}

a:link {
	color: #0000FF;
	text-decoration: none;
}
a:visited {
	color: #0000CC;
	text-decoration: none;
}
a:hover {
	color: #00CC00;
	text-decoration: none;
}
a:active {
	color: #33FFFF;
	text-decoration: none;
}

-->
</style>
</head>
<body>
<form id="form1" name="form1" method="post" action="" runat="server">
  <span class="style2">共有<span class="style3"><%= DataSet1.RecordCount %></span>筆行程  </span>
  <input type="button" name="Submit2" value="回行事曆" onclick="return w_back()" />
  <asp:Button ID="Button2" runat="server" Text="彙出Excel檔" OnClick="Button2_Click" />  
<asp:DataGrid AllowCustomPaging="true" AllowPaging="true" 
  AllowSorting="False" 
  AutoGenerateColumns="false" BackColor="#FF99cc" 
  CellPadding="3" 
  CellSpacing="0" 
  DataSource="<%# DataSet1.DefaultView %>" EditItemStyle-BackColor="#CCCC99" HeaderStyle-BackColor="#CCFF66" HeaderStyle-ForeColor="#FF0033" ID="DataGrid1" ItemStyle-BackColor="#FFCCCC" PagerStyle-Mode="NumericPages" PageSize="<%# DataSet1.PageSize %>" 
  runat="server" SelectedItemStyle-BackColor="#CCFFCC" SelectedItemStyle-ForeColor="" 
  ShowFooter="true" 
  ShowHeader="true" Width="100%" OnPageIndexChanged="DataSet1.OnDataGridPageIndexChanged" virtualitemcount="<%# DataSet1.RecordCount %>" 
>
    <headerstyle HorizontalAlign="left" BackColor="#FFCCFF" ForeColor="#3D3DB6" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Bold="true" Font-Size="smaller" />  
    <itemstyle BackColor="#FFCC99" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Size="smaller" />  
    <alternatingitemstyle BackColor="#E5E5E5" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Size="smaller" />  
    <footerstyle HorizontalAlign="left" BackColor="#E8EBFD" ForeColor="#3D3DB6" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Bold="true" Font-Size="smaller" />  
    <pagerstyle BackColor="white" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Size="smaller" />  
    <columns>
    <asp:TemplateColumn
	    HeaderText="啟始日期/時間" ItemStyle-Width="20%" 
        Visible="True">
      <ItemTemplate><%# showdate(DataSet1.FieldValue("m_date", Container)) %> <br />
        <%# showtime(DataSet1.FieldValue("m_time", Container)) %></ItemTemplate>
    </asp:TemplateColumn>
    <asp:TemplateColumn
	    HeaderText="行程" ItemStyle-Width="10%"
        Visible="True">
      <itemtemplate><%# showmessage(DataSet1.FieldValue("m_messeng", Container),DataSet1.FieldValue("m_hours", Container)) %> </itemtemplate>
    </asp:TemplateColumn>
                <asp:HyperLinkColumn
        DataNavigateUrlField="m_num"
        DataNavigateUrlFormatString="kc_cal_detail.aspx?m_num={0}"
        DataTextField="m_note" 
        Visible="True" target="cal_detail" 
        ItemStyle-Width="50%"
		HeaderText="主題"/><asp:TemplateColumn
	    HeaderText="結束日期/時間" ItemStyle-Width="20%" 
        Visible="True">
  <ItemTemplate><%# showdate(DataSet1.FieldValue("m_date_en", Container)) %> <br />
      <%# showtime(DataSet1.FieldValue("m_time_en", Container)) %></ItemTemplate>
</asp:TemplateColumn>
<asp:TemplateColumn HeaderText="修改" 
        Visible="True">
  <ItemTemplate>
    <input type="button" name="Submit3" value="修改" onclick= "MM_goToURL('self','kc_cal_update.aspx?m_num=<%# DataSet1.FieldValue("m_num", Container) %>');return document.MM_returnValue" />
  </ItemTemplate>
</asp:TemplateColumn>
</columns>

  </asp:DataGrid>
<asp:DataGrid
  AllowSorting="False" 
  AutoGenerateColumns="false" 
  CellPadding="3" 
  CellSpacing="0" 
  DataSource="<%# DataSet2.DefaultView %>" ID="DataGrid2" 
  runat="server" 
  ShowFooter="true" 
  ShowHeader="true" Visible="false" Width="100%" >
    <headerstyle HorizontalAlign="left" BackColor="#FFCCFF" ForeColor="#3D3DB6" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Bold="true" Font-Size="smaller" />  
    <footerstyle HorizontalAlign="left" BackColor="#E8EBFD" ForeColor="#3D3DB6" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Bold="true" Font-Size="smaller" />  
    <pagerstyle BackColor="white" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Size="smaller" />  
    <columns>
    <asp:TemplateColumn
	    HeaderText="啟始日期" ItemStyle-Width="10%" 
        Visible="True">
      <ItemTemplate><%# showdate(DataSet2.FieldValue("m_date", Container)) %> </ItemTemplate>
    </asp:TemplateColumn>
    <asp:TemplateColumn
	    HeaderText="時間" ItemStyle-Width="5%" 
        Visible="True">
      <ItemTemplate><%# showtime(DataSet2.FieldValue("m_time", Container)) %></ItemTemplate>
    </asp:TemplateColumn>

    <asp:TemplateColumn
	    HeaderText="行程" ItemStyle-Width="3%"
        Visible="True">
      <itemtemplate><%# DataSet2.FieldValue("m_messeng", Container) %> </itemtemplate>
    </asp:TemplateColumn>
    <asp:TemplateColumn
	    HeaderText="時數" ItemStyle-Width="3%"
        Visible="True">
      <itemtemplate><%# DataSet2.FieldValue("m_hours", Container) %> </itemtemplate>
    </asp:TemplateColumn>
                <asp:HyperLinkColumn
        DataNavigateUrlField="m_num"
        DataNavigateUrlFormatString="kc_cal_detail.aspx?m_num={0}"
        DataTextField="m_note" 
        Visible="True" target="cal_detail" 
        ItemStyle-Width="40%"
		HeaderText="主題"/>
		<asp:TemplateColumn
	    HeaderText="結束日期" ItemStyle-Width="10%" 
        Visible="True">
  <ItemTemplate><%# showdate(DataSet2.FieldValue("m_date_en", Container)) %> </ItemTemplate>
</asp:TemplateColumn>
		<asp:TemplateColumn
	    HeaderText="結束時間" ItemStyle-Width="5%" 
        Visible="True">
  <ItemTemplate><%# showtime(DataSet2.FieldValue("m_time_en", Container)) %></ItemTemplate>
</asp:TemplateColumn>
</columns>

  </asp:DataGrid>
</form>
<%
  session("cancel") = ""
  session("cancel") = request.url.tostring()
%>
</body>
</html>
