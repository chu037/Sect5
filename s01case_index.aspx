<%@ Page Language="VB" ContentType="text/html" ResponseEncoding="utf-8" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<%@ Register TagPrefix="MM" Namespace="DreamweaverCtrls" Assembly="DreamweaverCtrls,version=1.0.0.0,publicKeyToken=836f606ede05d46a,culture=neutral" %>
<MM:DataSet 
id="DataSet1"
runat="Server"
IsStoredProcedure="false"
ConnectionString='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_STRING_s01") %>'
DatabaseType='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_DATABASETYPE_s01") %>'
CommandText='<%# "SELECT distinct case_id, case_start, case_group, s01name, case_end, case_end_yes, case_yes, case_content, case_audit, admin_id, case_days, case_limit, case_alldays, case_result FROM s01_case  WHERE case_start between ? and ? AND case_group like ? AND s01name like ? AND case_end_yes like ? AND case_yes like ? AND case_content like ? AND case_audit like ?  AND case_result like ? ORDER BY case_start DESC" %>'
PageSize="50"
Debug="true"
><Parameters>
  <Parameter  Name="@case_start"  Value='<%# IIf((Request.QueryString("case_start") <> Nothing), Request.QueryString("case_start"), "#1/1/2006#") %>'  Type="Date"   />
   <Parameter  Name="@case_start1"  Value='<%# IIf((Request.QueryString("case_start1") <> Nothing), Request.QueryString("case_start1"), "#12/31/2099#") %>' Type="Date"   />
  
  <Parameter  Name="@case_group"  Value='<%# IIf((Request.QueryString("case_group") <> Nothing), Request.QueryString("case_group"), "%") %>'  Type="WChar"   />  
  <Parameter  Name="@s01name"  Value='<%# IIf((Request.QueryString("s01name") <> Nothing), Request.QueryString("s01name"), "%") %>'  Type="WChar"   />  
  <Parameter  Name="@case_end_yes"  Value='<%# IIf((Request.QueryString("case_end_yes") <> Nothing), Request.QueryString("case_end_yes"), "%") %>'  Type="WChar"   />  
  <Parameter  Name="@case_yes"  Value='<%# IIf((Request.QueryString("case_yes") <> Nothing), Request.QueryString("case_yes"), "%") %>'  Type="WChar"   />  
  <Parameter  Name="@case_content"  Value='<%# "%" + (IIf((Request.QueryString("case_content") <> Nothing), Request.QueryString("case_content"), "%")) + "%" %>'  Type="WChar"   />  
  <Parameter  Name="@case_audit"  Value='<%# IIf((Request.QueryString("case_audit") <> Nothing), Request.QueryString("case_audit"), "%") %>'  Type="WChar"   />  
  <Parameter  Name="@case_result"  Value='<%# "%" + (IIf((Request.QueryString("case_result") <> Nothing), Request.QueryString("case_result"), "%")) + "%" %>'  Type="WChar"   />  

</Parameters></MM:DataSet>
<MM:DataSet 
id="DataSet2"
runat="Server"
IsStoredProcedure="false"
ConnectionString='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_STRING_s01") %>'
DatabaseType='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_DATABASETYPE_s01") %>'
CommandText='<%# "SELECT *  FROM s01admin  ORDER BY admin_username ASC" %>'
Debug="true"
></MM:DataSet>
<MM:DataSet 
id="DataSet3"
runat="Server"
IsStoredProcedure="false"
ConnectionString='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_STRING_s01") %>'
DatabaseType='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_DATABASETYPE_s01") %>'
CommandText='<%# "SELECT distinct case_id, case_start, case_group, s01name, case_end, case_end_yes, case_yes, case_content, case_audit, admin_id, case_days, case_limit, case_alldays, case_result FROM s01_case  WHERE case_start between ? and ? AND case_group like ? AND s01name like ? AND case_end_yes like ? AND case_yes like ? AND case_content like ? AND case_audit like ? AND case_result like ? ORDER BY case_start DESC" %>'
Debug="true"
><Parameters>
  <Parameter  Name="@case_start"  Value='<%# IIf((Request.QueryString("case_start") <> Nothing), Request.QueryString("case_start"), "#1/1/2006#") %>'  Type="Date"   />
   <Parameter  Name="@case_start1"  Value='<%# IIf((Request.QueryString("case_start1") <> Nothing), Request.QueryString("case_start1"), "#12/31/2099#") %>' Type="Date"   />
  
  <Parameter  Name="@case_group"  Value='<%# IIf((Request.QueryString("case_group") <> Nothing), Request.QueryString("case_group"), "%") %>'  Type="WChar"   />  
  <Parameter  Name="@s01name"  Value='<%# IIf((Request.QueryString("s01name") <> Nothing), Request.QueryString("s01name"), "%") %>'  Type="WChar"   />  
  <Parameter  Name="@case_end_yes"  Value='<%# IIf((Request.QueryString("case_end_yes") <> Nothing), Request.QueryString("case_end_yes"), "%") %>'  Type="WChar"   />  
  <Parameter  Name="@case_yes"  Value='<%# IIf((Request.QueryString("case_yes") <> Nothing), Request.QueryString("case_yes"), "%") %>'  Type="WChar"   />  
  <Parameter  Name="@case_content"  Value='<%# "%" + (IIf((Request.QueryString("case_content") <> Nothing), Request.QueryString("case_content"), "%")) + "%" %>'  Type="WChar"   />  
  <Parameter  Name="@case_audit"  Value='<%# IIf((Request.QueryString("case_audit") <> Nothing), Request.QueryString("case_audit"), "%") %>'  Type="WChar"   />  
  <Parameter  Name="@case_result"  Value='<%# "%" + (IIf((Request.QueryString("case_result") <> Nothing), Request.QueryString("case_result"), "%")) + "%" %>'  Type="WChar"   />  

</Parameters></MM:DataSet>
<MM:PageBind runat="server" PostBackBind="false" />
<%@ Import Namespace="System.Data.OleDb" %>
<%@ Import Namespace="System.Globalization.Calendar" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<meta http-equiv="refresh" content="600"/>
<title>綜合行業科案件管制</title>

<script language="VB" runat="server">
   Sub Page_Load(sender As Object, e As EventArgs)
		If Session("MM_username")="" then
			Response.Write("您還沒有登入呢！，請點擊")
			Response.Write("<a href=s01login.aspx>登入</a>")
			Response.End()
		End If
     if not ispostback then
	 dim url_back
	  url_back = request.url.tostring()
	  session("cancel_case") = ""
	  session("cancel_case") = url_back
	  end if

        End Sub
   Sub starsearch(sender As Object, e As EventArgs) 
	  dim url0
	  url0 = "s01case_index.aspx?case_start=" & case_start.text & "&case_start1=" & case_start1.text & "&case_group=" & case_group.selecteditem.text & "&s01name=" & s01name.selecteditem.text & "&case_end_yes=" & case_end_yes.selecteditem.text & "&case_yes=" & case_yes.selecteditem.text & "&case_content=" & case_content.text & "&case_audit=" & case_audit.selecteditem.text & "&case_result=" & case_result.text 
	  response.Redirect( url0 ) ' 使用Server.Transfer亦可
   End Sub

   Sub starsearch_all(sender As Object, e As EventArgs) 
	  dim url0
	  url0 = "s01case_index.aspx?case_start=" & "" & "&case_start1=" & "" & "&case_group=" & "" & "&s01name=" & "" & "&case_end_yes=" & "" & "&case_yes=" & "" & "&case_content=" & "" & "&case_audit=" & "" & "&case_result=" & "" 
	  response.Redirect( url0 ) ' 使用Server.Transfer亦可
   End Sub


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
 function showyes(y)
 if y = "True" then
 showyes = "是"
 else
 showyes = "否"
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
        DataGrid3.visible = true
		DisableControls(DataGrid3)
        DataGrid3.RenderControl(ExportHtmlTextWriter)

        Response.AppendHeader("Content-Disposition", "attachment;filename=" + Date.Today.ToString("yyyy-MM-dd") + ".xls")
 'xlsApp = CreateObject("Excel.Application")   
 'xlsApp.Workbooks.Open("C:\Book1.xls")   

        Response.ContentEncoding = System.Text.Encoding.GetEncoding("big5")
        Response.ContentType = "application/vnd.ms-excel"
        Response.Write(ExportWriter.ToString())
        DataGrid3.visible = False
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

    </script>
<script type="text/JavaScript">
<!--
function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}

function MM_goToURL() { //v3.0
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"'");
}
function DelTitle(s01id) {//確認是否要刪除該記錄
	if (confirm('您確實要刪除該主題嗎？' + '\r\r' +
		'注意：如果您刪除該主題，' + '\r' +
		'哪麼該主題下所有的資料' + '\r' +
		'就將全部被刪除！')){
	window.location ='s01case_del.aspx?case_id=' + s01id;
    }
}
</script>
</head>

<body bgcolor="#FFFF66">

    <h3><strong><font color="#CC3300" face="Verdana, 新細明體">綜合行業科列管案件進度管制表</font></strong></h3>

    <form runat=server>

      <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="32%" valign="top"><table width="96%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
              </tr>
              <tr>
                <td width="28%">案件日期：</td>
                <td width="72%"><asp:TextBox ID="case_start" runat="server" Width="60" />~<asp:TextBox ID="case_start1" runat="server" Width="60" />
                <br />
                <font color="#990000">格式:YYYY/M/D</font></td>
              </tr>
              <tr>
                <td>案件類型：</td>
                <td><asp:DropDownList ID="case_group" runat="server">
                  <asp:ListItem></asp:ListItem>
      <asp:ListItem>申訴</asp:ListItem>
      <asp:ListItem>重大職災</asp:ListItem>
      <asp:ListItem>重大職災(其他)</asp:ListItem>
      <asp:ListItem>重大職災(勞安法以外)</asp:ListItem>
      <asp:ListItem>重大職災(認定中)</asp:ListItem>
      <asp:ListItem>非重大職災</asp:ListItem>
      <asp:ListItem>專案</asp:ListItem>
      <asp:ListItem>其他</asp:ListItem>
                </asp:DropDownList></td>
              </tr>
              <tr>
                <td>承辦人：</td>
                <td><asp:DropDownList ID="s01name" runat="server" DataSource="<%# dataset2.defaultview %>" DataTextField="s01name" DataValueField="s01name">

                </asp:DropDownList></td>
              </tr>
              <tr>
                <td>已結案：</td>
                <td><asp:DropDownList ID="case_end_yes" runat="server">
                  <asp:ListItem></asp:ListItem>
                  <asp:ListItem>是</asp:ListItem>
                  <asp:ListItem>否</asp:ListItem>
                </asp:DropDownList></td>
              </tr>
              <tr>
                <td>已登錄：</td>
                <td><asp:DropDownList ID="case_yes" runat="server">
                  <asp:ListItem></asp:ListItem>
                  <asp:ListItem>是</asp:ListItem>
                  <asp:ListItem>否</asp:ListItem>
                  <asp:ListItem>不需</asp:ListItem>
                </asp:DropDownList></td>
              </tr>
              <tr>
                <td>內容：</td>
                <td><asp:TextBox ID="case_content" runat="server" Width="100" /></td>
              </tr>
              <tr>
                <td>辦理情形：</td>
                <td><asp:TextBox ID="case_result" runat="server" Width="100" /></td>
              </tr>

              <tr>
                <td>逾期：</td>
                <td><asp:DropDownList ID="case_audit" runat="server">
                  <asp:ListItem></asp:ListItem>
                  <asp:ListItem>逾期</asp:ListItem>
                </asp:DropDownList>
                  <font color="#FF0000">                  (逾期尚未結案)</font></td>
              </tr>
            </table>
              <p>
              <asp:Button ID="case_search" runat="server" Text="查詢" OnClick="starsearch" />                                            
              <asp:Button ID="case_all" runat="server" Text="顯示所有案件" OnClick="starsearch_all" />                            </p>
              <p>&nbsp;</p>
              <p><br />
                <br />
                <br />
                <br />
                <br />
                <br />
              </p>
              <p>&nbsp;</p></td>
            <td width="68%" colspan="" valign="top">
                共<strong><font color="#CC3300"><%= DataSet1.RecordCount %></font></strong>筆<strong><font color="#CC3300">(統計資料自2012/5/1起)</font></strong><font color="#CC3300"><font color="#000000">
                <asp:Button ID="Button2" runat="server" Text="匯出Excel檔" OnClick="Button2_Click" />      

                </font>(<font color="#0000FF">請點選&quot;修改&quot;上傳相關資料結案</font>)</font>
                <asp:DataGrid 
  AllowCustomPaging="true" 
  AllowPaging="true" 
  AllowSorting="true" 
  AutoGenerateColumns="false" 
  CellPadding="3" 
  CellSpacing="0" 
  DataSource="<%# DataSet1.DefaultView %>" id="DataGrid1" 
  PagerStyle-Mode="NumericPages" 
  PageSize="<%# DataSet1.PageSize %>" 
  runat="server" 
  ShowFooter="false" 
  ShowHeader="true" 
  OnPageIndexChanged="DataSet1.OnDataGridPageIndexChanged"  
  VirtualItemCount="<%# DataSet1.RecordCount %>" 
>
                  <HeaderStyle HorizontalAlign="center" BackColor="#E8EBFD" ForeColor="#3D3DB6" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Bold="true" Font-Size="smaller" />                
                  <ItemStyle BackColor="#F2F2F2" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Size="smaller" />                
                  <AlternatingItemStyle BackColor="#E5E5E5" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Size="smaller" />                
                  <FooterStyle HorizontalAlign="center" BackColor="#E8EBFD" ForeColor="#3D3DB6" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Bold="true" Font-Size="smaller" />                
                  <PagerStyle BackColor="white" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Size="smaller" />                
                  <Columns>
                  <asp:BoundColumn DataField="case_id" 
        HeaderText="case_id" 
        ReadOnly="true" 
        Visible="false"/>                  
                  <asp:BoundColumn DataField="case_group" 
        HeaderText="類型" 
        ReadOnly="true" 
        Visible="True" ItemStyle-Width="12%" HeaderStyle-HorizontalAlign="left"/>                  
                  <asp:TemplateColumn HeaderText="案件日期" 
        Visible="True" ItemStyle-Width="10%" HeaderStyle-HorizontalAlign="left">
                    <ItemTemplate><%# showdate(DataSet1.FieldValue("case_start", Container)) %> </ItemTemplate>
                  </asp:TemplateColumn >
                  <asp:BoundColumn DataField="s01name" 
        HeaderText="承辦人" 
        ReadOnly="true" 
        Visible="True" ItemStyle-Width="10%" HeaderStyle-HorizontalAlign="left"/>                  
<asp:HyperLinkColumn
        DataNavigateUrlField="case_id"
        DataNavigateUrlFormatString="s01case_detail.aspx?case_id={0}"
        DataTextField="case_content" 
        Visible="True" target="case_detail" 
        HeaderText="案件內容"/>                  
<asp:BoundColumn DataField="case_days" 
        HeaderText="已辦天數" 
        ReadOnly="true" 
        Visible="True" ItemStyle-Width="10%" HeaderStyle-HorizontalAlign="left"/>
<asp:BoundColumn DataField="case_limit" 
        HeaderText="限辦天數" 
        ReadOnly="true" 
        Visible="True" ItemStyle-Width="10%" HeaderStyle-HorizontalAlign="left"/>
<asp:BoundColumn DataField="case_audit" 
        HeaderText="逾辦" 
        ReadOnly="true" 
        Visible="True" ItemStyle-ForeColor="#CC3300" ItemStyle-Width="8%" HeaderStyle-HorizontalAlign="left"/>
<asp:TemplateColumn HeaderText="結案日期" 
        Visible="True" ItemStyle-Width="10%" HeaderStyle-HorizontalAlign="left">
  <ItemTemplate><%# showdate(DataSet1.FieldValue("case_end", Container)) %> </ItemTemplate>
</asp:TemplateColumn>
<asp:BoundColumn DataField="case_end_yes" 
        HeaderText="已結案" 
        ReadOnly="true" 
        Visible="True" ItemStyle-Width="8%" HeaderStyle-HorizontalAlign="left"/>
<asp:BoundColumn DataField="case_alldays" 
        HeaderText="完成天數" 
        ReadOnly="true" 
        Visible="false"/>
<asp:TemplateColumn HeaderText="修改" 
        Visible="true">
  <ItemTemplate>
    <input name="Submit2" type="button" onclick="MM_goToURL('parent','s01case_update_normal.aspx?case_id=<%# DataSet1.FieldValue("case_id", Container) %>');return document.MM_returnValue" value="修改" />
    <input name="Submit3" type="hidden" onClick="DelTitle('<%# DataSet1.FieldValue("case_id", Container) %>')" value="刪除" />
  </ItemTemplate>
</asp:TemplateColumn>
                  </Columns>
              </asp:DataGrid>
              <p>
 <asp:DataGrid 
  AllowPaging="false" 
  AllowSorting="false" 
  AutoGenerateColumns="false" 
  CellPadding="3" 
  CellSpacing="0" 
  DataSource="<%# DataSet3.DefaultView %>" id="DataGrid3" 
  runat="server" 
  ShowFooter="false" 
  ShowHeader="true" Visible="false" 
>
                  <HeaderStyle HorizontalAlign="center" BackColor="#E8EBFD" ForeColor="#3D3DB6" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Bold="true" Font-Size="smaller" />                
                  <ItemStyle BackColor="#F2F2F2" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Size="smaller" />                
                  <AlternatingItemStyle BackColor="#E5E5E5" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Size="smaller" />                
                  <FooterStyle HorizontalAlign="center" BackColor="#E8EBFD" ForeColor="#3D3DB6" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Bold="true" Font-Size="smaller" />                
                  <PagerStyle BackColor="white" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Size="smaller" />                
                  <Columns>
                  <asp:BoundColumn DataField="case_id" 
        HeaderText="case_id" 
        ReadOnly="true" 
        Visible="false"/>                  
                  <asp:BoundColumn DataField="case_group" 
        HeaderText="類型" 
        ReadOnly="true" 
        Visible="True" ItemStyle-Width="12%" HeaderStyle-HorizontalAlign="left"/>                  
                  <asp:TemplateColumn HeaderText="案件日期" 
        Visible="True" ItemStyle-Width="10%" HeaderStyle-HorizontalAlign="left">
                    <ItemTemplate><%# showdate(DataSet3.FieldValue("case_start", Container)) %> </ItemTemplate>
                  </asp:TemplateColumn >
                  <asp:BoundColumn DataField="s01name" 
        HeaderText="承辦人" 
        ReadOnly="true" 
        Visible="True" ItemStyle-Width="10%" HeaderStyle-HorizontalAlign="left"/>                  
                  <asp:BoundColumn DataField="case_content" 
        HeaderText="案件內容" 
        ReadOnly="true" 
        Visible="True" ItemStyle-Width="20%" HeaderStyle-HorizontalAlign="left"/>                  
                  <asp:BoundColumn DataField="case_result" 
        HeaderText="辦理情形" 
        ReadOnly="true" 
        Visible="True" ItemStyle-Width="20%" HeaderStyle-HorizontalAlign="left"/>                  
                  <asp:BoundColumn DataField="case_days" 
        HeaderText="已辦天數" 
        ReadOnly="true" 
        Visible="True" ItemStyle-Width="10%" HeaderStyle-HorizontalAlign="left"/>                  
                  <asp:BoundColumn DataField="case_limit" 
        HeaderText="限辦天數" 
        ReadOnly="true" 
        Visible="True" ItemStyle-Width="10%" HeaderStyle-HorizontalAlign="left"/>                  
                  <asp:BoundColumn DataField="case_audit" 
        HeaderText="逾辦" 
        ReadOnly="true" 
        Visible="True" ItemStyle-ForeColor="#CC3300" ItemStyle-Width="8%" HeaderStyle-HorizontalAlign="left"/>                  
                  <asp:TemplateColumn HeaderText="結案日期" 
        Visible="True" ItemStyle-Width="10%" HeaderStyle-HorizontalAlign="left">
                    <ItemTemplate><%# showdate(DataSet3.FieldValue("case_end", Container)) %> </ItemTemplate>
                  </asp:TemplateColumn>
                  <asp:BoundColumn DataField="case_end_yes" 
        HeaderText="已結案" 
        ReadOnly="true" 
        Visible="True" ItemStyle-Width="8%" HeaderStyle-HorizontalAlign="left"/>                  
                  <asp:BoundColumn DataField="case_alldays" 
        HeaderText="完成天數" 
        ReadOnly="true" 
        Visible="false"/>                  
<asp:TemplateColumn HeaderText="修改" 
        Visible="false">
                  </asp:TemplateColumn>
                  </Columns>
              </asp:DataGrid>
      </table>
        <p>
        <asp:Label id=Label1 runat="server" />
        
</form>

</body>
</html>