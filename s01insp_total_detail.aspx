<%@ Page Language="VB" ContentType="text/html"  %>
<%@ Register TagPrefix="MM" Namespace="DreamweaverCtrls" Assembly="DreamweaverCtrls,version=1.0.0.0,publicKeyToken=836f606ede05d46a,culture=neutral" %>
<MM:DataSet 
id="DataSet1"
runat="Server"
IsStoredProcedure="false"
ConnectionString='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_STRING_s01") %>'
DatabaseType='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_DATABASETYPE_s01") %>'
CommandText='<%# "SELECT * FROM s01_total WHERE t_id = ?" %>'
Debug="true"
>
  <Parameters>
    <Parameter  Name="@t_id"  Value='<%# IIf((Request.QueryString("t_id") <> Nothing), Request.QueryString("t_id"), "") %>'  Type="Integer"   />
  </Parameters>
</MM:DataSet>
<MM:PageBind runat="server" PostBackBind="true" />

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<script runat="server">
	Sub Page_Load(Src As Object, E As EventArgs)
		'If Session("MM_username")="" then
			'Response.Write("您還沒有登入呢！，請點擊")
			'Response.Write("<a href=s01insp_list_login.aspx>登入</a>")
			'Response.End()
		'End If
		if request("pho")="" then
		 Button6.enabled = "true"
		end if 
	End Sub

   Sub starsearch(sender As Object, e As EventArgs) 
	  dim url
	  url = "s01insp_t_update.aspx?t_id=" & label1.text
	  Response.Redirect( url ) ' 使用Server.Transfer亦可
   End Sub
   Sub staradd(sender As Object, e As EventArgs) 
	  dim url2
	  url2 = "s01insp_add0.aspx?p_num=" & label1.text & "&p_name=" & text3.text
	  Response.Redirect( url2 ) ' 使用Server.Transfer亦可
   End Sub
  dim url3
   Sub photo_s(sender As Object, e As EventArgs) 
	  url3 = "coaching_plant_index.aspx?p_id=" & label3.text
   End Sub

  dim map as string
  dim x01
   Sub map_search(sender As Object, e As EventArgs) 
    
	  x01 = text1.text
	  map = "http://maps.google.com.tw/maps?hl=zh-TW&q=" & x01
   End Sub
  
  dim map2 as string
  dim x02
   Sub map_search2(sender As Object, e As EventArgs) 
    
	  x02 = text11.text
	  map2 = "http://maps.google.com.tw/maps?hl=zh-TW&q=" & x02
   End Sub

</script>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<title>事業單位基本資料</title>

<style type="text/css">
<!--
.style1 {color: #CC0000}
.style2 {color: #CC3300}
.style3 {color: #000000}
.style4 {color: #0000FF}
-->
</style>
<script type="text/JavaScript">
<!--
function MM_goToURL() { //v3.0
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"'");
}
//-->
</script>
</head>
<body >	<!--顯示分行結果須用下面程式 -->
<script runat="server">
  Function Clean(str)
   str=Replace(str, vbCrLf, "<br>")
   Clean=Replace(str, chr(32), "&nbsp;&nbsp;")
  End Function
 </script> 
<form runat="server">
  <asp:Button ID="Button1" runat="server" Text="修改基本資料" OnClick="starsearch" />
<asp:Button Enabled="false" ID="Button6" runat="server" Text="圖照系統" OnClick="photo_s" />
<span class="style1">  (※地圖為找路參考用)</span>  
  
  <asp:Label ID="Label1" runat="server" Text='<%# DataSet1.FieldValue("t_id", Container) %>' Visible="false" />
  <asp:Label ID="Label2" runat="server" Text='<%# DataSet1.FieldValue("t_address", Container) %>' Visible="false" />
   <asp:Label ID="Label3" runat="server" Text='<%# DataSet1.FieldValue("t_insu", Container) %>' Visible="false" />
 
  <asp:TextBox ID="text1" ReadOnly="true" runat="server" text='<%# DataSet1.FieldValue("t_address", Container) %>' Visible="false" />
  <asp:TextBox ID="text2" ReadOnly="true" runat="server" text='<%# DataSet1.FieldValue("t_id", Container) %>' Visible="false" />
  <asp:TextBox ID="text3" ReadOnly="true" runat="server" text='<%# DataSet1.FieldValue("t_name", Container) %>' Visible="false" />
<asp:TextBox ID="text11" ReadOnly="true" runat="server" text='<%# DataSet1.FieldValue("t_address_m", Container) %>' Visible="false" />
  <table width="100%" border="1" cellspacing="0" cellpadding="0">
  <tr>
    <td width="12%" bgcolor="#FFCCFF">單位名稱：</td>
    <td width="53%" bgcolor="#CCFFFF"><%# DataSet1.FieldValue("t_name", Container) %></td>
    <td width="13%" bgcolor="#FFCCFF">負責人：</td>
    <td width="22%" bgcolor="#CCFFFF"><%# DataSet1.FieldValue("t_boss", Container) %></td>
  </tr>
  <tr>
    <td bgcolor="#FFCCFF">地址：</td>
    <td bgcolor="#CCFFFF"><%# DataSet1.FieldValue("t_address", Container) %><asp:Button ID="Button2" runat="server" Text="地圖" OnClick="map_search" /></td>
    <td bgcolor="#FFCCFF">電話：</td>
    <td bgcolor="#CCFFFF"><%# DataSet1.FieldValue("t_tel", Container) %></td>
  </tr>
  <tr>
    <td bgcolor="#FFCCFF">地址：</td>
    <td bgcolor="#CCFFFF"><%# DataSet1.FieldValue("t_address_m", Container) %><asp:Button ID="Button21" runat="server" Text="地圖" OnClick="map_search2" /></td>
    <td bgcolor="#FFCCFF">電話：</td>
    <td bgcolor="#CCFFFF"><%# DataSet1.FieldValue("t_tel", Container) %></td>
  </tr>
  <tr>
    <td bgcolor="#FFCCFF">行業別：</td>
    <td bgcolor="#CCFFFF"><%# DataSet1.FieldValue("t_group.t_g_name", Container) %></td>
    <td bgcolor="#FFCCFF">傳真：</td>
    <td bgcolor="#CCFFFF">&nbsp;</td>
  </tr>
  <tr>
    <td bgcolor="#FFCCFF">e-mail：</td>
    <td bgcolor="#CCFFFF"><%# DataSet1.FieldValue("t_email", Container) %></td>
    <td bgcolor="#FFCCFF">勞工人數：</td>
    <td bgcolor="#CCFFFF"><%# DataSet1.FieldValue("t_per", Container) %></td>
  </tr>

  <tr>
    <td bgcolor="#FFCCFF">受檢場所：</td>
    <td bgcolor="#CCFFFF"><%# clean(DataSet1.FieldValue("t_place", Container)) %></td>
    <td bgcolor="#FFCCFF">連絡人：</td>
    <td bgcolor="#CCFFFF"><%# DataSet1.FieldValue("t_sh", Container) %></td>
  </tr>
  <tr>
    <td bgcolor="#FFCCFF">相關危害：</td>
    <td bgcolor="#CCFFFF"><%# clean(DataSet1.FieldValue("t_hazard", Container)) %></td>
    <td bgcolor="#FFCCFF">手機：</td>
    <td bgcolor="#CCFFFF"><%# DataSet1.FieldValue("t_shcell", Container) %></td>
  </tr>
</table>
</form>

<p>&nbsp;</p>
<%
  session("cancel_insp") = ""
  session("cancel_insp") = request.url.tostring()
%>

</body>
</html>

<% 
If x01 <>"" Then
%>

<script language = "JavaScript">
var p;
p = "<%= map %>";
window.open(p);

</Script>
<% 
x01 =""
End If
%>

<% 
If x02 <>"" Then
%>

<script language = "JavaScript">
var p;
p = "<%= map2 %>";
window.open(p);

</Script>
<%
x02=""
End If
%>
