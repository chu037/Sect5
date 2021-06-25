<%@ Page Language="VB" ContentType="text/html" ResponseEncoding="utf-8" %>
<%@ Register TagPrefix="MM" Namespace="DreamweaverCtrls" Assembly="DreamweaverCtrls,version=1.0.0.0,publicKeyToken=836f606ede05d46a,culture=neutral" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<title>新增案件</title>
<script language="vb" runat="server">
  Sub page_load(sender As Object, e As EventArgs)
	msg_001.text = ""
		If Session("MM_chief_num")="" then
			Response.Write("您還沒有登入呢！，請點擊")
			Response.Write("<a href=s01dellogin.aspx>登入</a>")
			Response.End()
		End If

  if not Ispostback Then
    case_start.text = request("case_start")
	if case_start.text = "" then
      response.Redirect("s01case_admin.aspx")
    else
	back.text = session("cancel_case_admin")
	 end if
	end if
  end sub

   Sub addchk(sender As Object, e As EventArgs) 
 if case_content.text = "" then
  msg_001.text = "請先輸入事業單位名稱"
  exit sub
 else  
	  dim url0
	  dim asc_m as string
	  if case_content.text <> "" then
	   asc_m = case_content.text
	  else
	   asc_m = "" 
	  end if
	  url0 = "s01case_add_chk.aspx?case_start=" & case_start.text & "&case_content=" & trans(case_content.text)
	  response.Redirect( url0 ) ' 使用Server.Transfer亦可
    end if
   End Sub

function trans(str01)
 if str01 <> ""
  dim i as integer
  dim stra
  stra = ""
   for i = 1 to len(str01)-1
	stra &= asc(mid(str01,i,1)) & ","
   next
    if i = len(str01)
	stra = stra & asc(mid(str01,i,1))
	end if
   trans = stra
  end if 
   end function	

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
Sub startcancel(sender As Object, e As EventArgs) 
 response.Redirect(back.text)
end sub
</script>


<style type="text/css">
<!--
.style1 {
	font-size: 16px;
	font-weight: bold;
}
.style2 {
	color: #FF0000;
	font-weight: bold;
}
.style3 {	color: #CC3300;
	font-weight: bold;
	font-size: 16px;
}
-->
</style>
</head>
<body>
<p class="style1">新增案件：</p>
<form runat="server" name='form_add' id="form_add">
<table width="529" border="0" cellspacing="1" cellpadding="0">
  <tr>
    <td width="80" bgcolor="#CCFFFF">案件日期：</td>
    <td width="446" bgcolor="#CCCC66"><asp:TextBox ID="case_start" ReadOnly="true" runat="server" Wrap="false" /></td>
  </tr>
  <tr>
    <td width="80" bgcolor="#CCFFFF">案件內容：</td>
    <td width="446" bgcolor="#CCCC66">
    <asp:TextBox ID="case_content" runat="server" />    
    請先輸入案件內容(部分名稱即可)</td>
  </tr>

</table>
  <asp:Button ID="Button1" runat="server" Text="查詢" OnClick="addchk" />
  <span class="style1">
  <input type="reset" name="Submit2" value="重新填寫" />
  </span>
  <span class="style3">
  <asp:Button ID="Buttonc" runat="server" Text="取消" OnClick='startcancel' />  </span>
  <asp:TextBox ID="back" runat="server" Visible="false" Columns="50" />
  
  <asp:Label ForeColor="#CC0000" ID="msg_001" runat="server" />
  
  <p>
  <HR>
</form>
</body>
</html>