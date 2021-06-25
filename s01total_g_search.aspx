<%@ Page Language="VB" ContentType="text/html" %><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>行業查詢</title>
<style type="text/css">
<!--
.style1 {color: #000000}
.style2 {color: #FF8000}
.style3 {
	color: #CC6633;
	font-weight: bold;
}
-->
</style>

</head>
<body>
<form action="" method="post" name="form1" id="form1" runat="server">
  <p><span class="style3">行業別</span><span class="style1">-查詢</span>：請輸入篩選條件後按&quot;<span class="style1">查詢</span>&quot;</p>
  <table width="100%" border="0" cellspacing="0" cellpadding="0">

    <tr>
      <td>行業名稱：</td>
      <td><asp:TextBox ID="t_g_name" runat="server" Columns="30" />

        <br />
        (可輸入行業別部分名稱)</td>
    </tr>
    <tr>
      <td>行業代碼：</td>
      <td><asp:TextBox ID="t_g_num" runat="server" Columns="10" /></td>
	</tr>

  </table>
  <br>
    <asp:Button ID="insp_search" runat="server" Text="查詢" OnClick="starsearch" />
    <input type="reset" name="Submit" value="重新填寫" />
    <input name="Submit2" type="button" onclick="window.close()" value="取消" />
</form>

</body>
</html>
<script Language="VB" runat="server">
   Sub page_load(sender As Object, e As EventArgs) 
   End Sub
   
   Sub starsearch(sender As Object, e As EventArgs) 
	  dim url
	  url = "s01total_g_result.aspx?t_g_name=" & t_g_name.text & "&t_g_num=" & t_g_num.text
	  response.Redirect( url ) ' 使用Server.Transfer亦可
   End Sub
   Sub plant_search(sender As Object, e As EventArgs) 
	  dim url
	  url = "s01insp_plant_search.aspx"
	  response.Redirect( url ) ' 使用Server.Transfer亦可
   End Sub
   Sub list_search(sender As Object, e As EventArgs) 
	  dim url
	  url = "s01insp_list.aspx"
	  response.Redirect( url ) ' 使用Server.Transfer亦可
   End Sub

</script>
<script language = "JavaScript">
<!--

window.opener.document.getElementById("t_g_1_t").value = "%";
window.opener.document.getElementById("t_g_2_t").value = "%";
window.opener.document.getElementById("t_g_3_t").value = "%";
window.opener.document.getElementById("t_g_num_t").value = "%";
window.opener.document.getElementById("t_g_id_t").value = "";
<!--window.opener.document.location.reload();
//-->

</script>