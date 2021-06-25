<%@ Page Language="VB" ContentType="text/html" %>
<%@ Register TagPrefix="MM" Namespace="DreamweaverCtrls" Assembly="DreamweaverCtrls,version=1.0.0.0,publicKeyToken=836f606ede05d46a,culture=neutral" %>
<MM:PageBind runat="server" PostBackBind="true" />
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<script runat="server">
	Sub Page_Load(Src As Object, E As EventArgs)

		If Session("MM_chief_num")="" then
			Response.Write("您還沒有登入呢！，請點擊")
			Response.Write("<a href=s01dellogin.aspx>登入</a>")
			Response.End()
		End If
	End Sub
</script>
<html>
<head>
  <meta http-equiv="Content-Language" content="zh-tw">
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <meta name="GENERATOR" content="Microsoft FrontPage 4.0">
  <meta name="ProgId" content="FrontPage.Editor.Document">
  <title>點選要修改的訊息</title>
  <style type="text/css">
<!--
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

body {
	font-family: "新細明體";
	font-size: 14px;
	font-style: normal;
	background-color: #FFFFCC;
}

.style2 {
	color: #993333;
	font-size: 24px;
	font-weight: bold;
}
.style5 {font-size: 16px; color: #990033; }
.style9 {
	font-size: 12px;
	color: #000000;
}
.style12 {font-size: 12px; color: #000000}
.style13 {font-size: 14px}
.style14 {color: #000099}
-->
  </style>
<script language="JavaScript" type="text/javascript">
<!--

function DelTitle(s01id) {//確認是否要刪除該記錄
	if (confirm('您確實要刪除該主題嗎？' + '\r\r' +
		'注意：如果您刪除該主題，' + '\r' +
		'哪麼該主題下所有的資料' + '\r' +
		'就將全部被刪除！')){
	window.location ='s01delprocess.aspx?s01id=' + s01id;
    }
}
//-->
</script>
</head>
<body>
<script language="VB" runat="server">
 function shownew(s01time)
 Dim time1 = s01time
 Dim time2 = Now
 Dim diff = DateDiff("d",time1,time2)  
 if diff < 7
 shownew = "<img border=0 src='images/new5.gif'>" 
 end if
 end function
 function showpoint(s01data)
 if s01data <> ""
 showpoint = "<img border=0 src='images/point01.gif'>" 
 end if
 end function
</script>
	  
<div align="center" class="style2">
  <p class="style14">請點選要刪改的訊息</p>
  <p align="left"><span class="style13"><a href="s01logout.aspx" target="_self">返回最新訊息</a></span></p>
</div>
<form runat="server">
  <table width="100%" border="1" cellpadding="0" cellspacing="0" bgcolor="#FFCCCC">
    <tr>
      <td colspan="6">
          訊息主題搜尋：
    <input name="filter" type="text" id="filter" size="15" /><input type="submit" name="Submit" value="搜尋">
    </div> <a href="s01cal_del.aspx">刪除行事曆</a>    
    ---<a href="s01member_index.aspx" target="_blank">組員管理</a>
    ---<a href="s01case_admin.aspx" target="_blank">案件管制</a>
    ---<a href="s01insp_total_s_index.aspx" target="_blank">轄區管理</a>

    </tr>

  </table>
</form>
</body>
</html>
