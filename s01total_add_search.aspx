<%@ Page Language="VB" ContentType="text/html" %>
<%@ Register TagPrefix="MM" Namespace="DreamweaverCtrls" Assembly="DreamweaverCtrls,version=1.0.0.0,publicKeyToken=836f606ede05d46a,culture=neutral" %>
<MM:DataSet 
id="DataSet1"
runat="Server"
IsStoredProcedure="false"
ConnectionString='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_STRING_s01") %>'
DatabaseType='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_DATABASETYPE_s01") %>'
CommandText='<%# "SELECT * FROM s01total WHERE t_keyin <> ? ORDER BY t_update_time DESC" %>'
PageSize="50"
Debug="true"
><Parameters>
<Parameter  Name="@t_keyin"  Value='<%# "t_keyin" %>'  Type="WChar"   /></Parameters></MM:DataSet>
<MM:PageBind runat="server" PostBackBind="true" />

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<title>新增事業單位查詢</title>
<style type="text/css">
<!--
.style1 {color: #000000}
.style3 {
	color: #cc6633;
	font-weight: bold;
}
.style4 {font-size: 18px}
body {
	background-color: #CCFFCC;
}
.style5 {color: #CC00CC}
.style7 {
	font-size: 18px;
	color: #990000;
	font-weight: bold;
}
.style8 {color: #3300FF}
-->
</style>
<script Language="VB" runat="server">
   Sub page_load(sender As Object, e As EventArgs) 
		If Session("MM_username")="" then
			Response.Write("您還沒有登入呢！，請點擊")
			Response.Write("<a href=s01insp_list_login.aspx>登入</a>")
			Response.End()
		End If
if ispostback then
exit sub
end if
   End Sub
   Sub starsearch(sender As Object, e As EventArgs) 
	  dim url
	  dim p_num
	  url = "s01insp_plant_numcheck.aspx?p_num=" & p_num.text 
	  Response.Redirect( url ) ' 使用Server.Transfer亦可
   End Sub
 function showdate(time01)
  if time01 <> ""
  showdate = FormatDateTime(time01, DateFormat.ShortDate) &"-"& FormatDateTime(time01, DateFormat.ShortTime)
 end if
 end function
</script>
<script type="text/JavaScript">
<!--
function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}

function MM_popupMsg(msg) { //v1.0
  alert(msg);
}
function Mcheck(){
  var i;
  var pnum = document.form1.p_num.value;
  var pnum05
  var pnum01 = pnum.substring(0,3);
  var pnum02 = pnum.substring(8,9);
  var pnum03 = pnum.substring(3,9);
  var pnum04 = pnum.substring(9,20);
  var pnum06 = pnum.substring(3,20);
  var pnum07 = pnum.substring(0,2);
  var pnum08 = pnum.substring(2,8);
  var pnum09 = pnum.substring(9,10);
  var pnum10 ;
if (pnum07 =="t-") {
  for (i=2; i<=8; i++) { 
  var pnum10 = pnum.substring(i,i+1);
 	if (isNaN(pnum10) | pnum10 ==" " | pnum10 =="　" | pnum10 =="") {
        window.alert("臨時工廠證號請輸入t-XXXXXXX，t-及7碼數字");
		return false }
	if (pnum09 !="") {
        window.alert("t-後數字超過7碼，臨時工廠證號請輸入t-XXXXXXX");
		return false }
}
     return true;

}
	if (pnum01 =="un-" && pnum06 != "") {
        window.alert("未登記工廠證號請輸入un-即可");
		return false }
	if (pnum01 =="un-" && pnum06 =="") {
		return true }
  for (i=3; i<=9; i++) { 
  var pnum05 = pnum.substring(i,i+1);
	if (isNaN(pnum03) | pnum05 ==" " | pnum05 =="　" ) {
        window.alert(pnum05 + ",含有非數字或空格,正確格式：64-XXXXXX、99-XXXXXX或12-XXXXXX");
        return false }}
    if (pnum01 !="64-" && pnum01 !="99-" && pnum01 !="12-") {
        window.alert(pnum01 + ",未以64-、99-、12-或un-開頭");
        return false }
    if (pnum02 =="") {
        window.alert(pnum03 + ",-後不足6碼,正確格式：64-XXXXXX、99-XXXXXX或12-XXXXXX");
        return false }
    if (isNaN(pnum03)) {
        window.alert(pnum03 + ",含有非數字,正確格式：64-XXXXXX、99-XXXXXX或12-XXXXXX");
        return false }
    if (pnum04 !="") {
        window.alert(pnum03 + ",-後超過6碼,正確格式：64-XXXXXX、99-XXXXXX或12-XXXXXX");
        return false }
     return true;

}

function Check_prenum(idvalue) {
   var tmp = new String("12121241");
   var sum = 0;
   re = /^\d{8}$/;
   if (!re.test(idvalue)) {
       alert("請輸入8碼數字！");
       return false;
    }

   for (i=0; i< 8; i++) {
     s1 = parseInt(idvalue.substr(i,1));
     s2 = parseInt(tmp.substr(i,1));
     sum += cal(s1*s2);
   }
 
   if (!valid(sum)) {
      if (idvalue.substr(6,1)=="7") {
	  if(!valid(sum+1)) {
       alert("格式不對！");
       return false;
    }	  
   return true;
   }  
       alert("格式不對！");
       return false;
  }
}

function valid(n) {
   return(n%10 == 0)?true:false;
   }

function cal(n) {
   var sum=0;
   while (n!=0) {
      sum += (n % 10);
      n = (n - n%10) / 10;  // 取整數
     }
   return sum;
}

function MM_goToURL() { //v3.0
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"'");
}
//-->
</script>
</head>
<body>
<form action="s01insp_total_numcheck.aspx" method="post" name="form1" id="form1" onSubmit="return Check_prenum(document.form1.p_num.value)">
  <p><span class="style5">新增事業單位</span>：請輸入統一編號後按&quot;<span class="style1">新增</span>&quot;</p>
  <hr />
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
 <tr>
      <td colspan="2">&nbsp;</td>
    </tr>
 <tr>
      <td colspan="2">&nbsp;</td>
    </tr>

 <tr>
 <tr>
      <td width="15%">統一編號：</td>
	  <td width="90%"><input name="p_num" type="text" id="p_num"  />
      (共8碼)</td>
    </tr>
      <td width="15%">&nbsp;</td>
	  <td width="90%">&nbsp;</td>
    </tr>

  </table>
  <p>&nbsp;
    <input type="submit" name="Submit" value="新增" />
    <input name="Submit2" type="submit" onclick="MM_goToURL('parent','s01total_search.aspx');return document.MM_returnValue" value="取消" />
  </p>
  <hr />
<p>最近更新資料：</p>
</form>
<form runat="server">
  <asp:DataGrid 
  AllowCustomPaging="true" 
  AllowPaging="true" 
  AllowSorting="False" 
  AutoGenerateColumns="false" 
  CellPadding="3" 
  CellSpacing="0" 
  DataSource="<%# DataSet1.DefaultView %>" ID="DataGrid1" 
  PagerStyle-Mode="NumericPages" 
  PageSize="<%# DataSet1.PageSize %>" 
  runat="server" 
  ShowFooter="false" 
  ShowHeader="true" Width="100%" 
  OnPageIndexChanged="DataSet1.OnDataGridPageIndexChanged" 
  virtualitemcount="<%# DataSet1.RecordCount %>" 
>
    <headerstyle HorizontalAlign="left" BackColor="#E8EBFD" ForeColor="#3D3DB6" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Bold="true" Font-Size="smaller" />  
    <itemstyle BackColor="#F2F2F2" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Size="smaller" />  
    <alternatingitemstyle BackColor="#E5E5E5" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Size="smaller" />  
    <footerstyle HorizontalAlign="center" BackColor="#E8EBFD" ForeColor="#3D3DB6" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Bold="true" Font-Size="smaller" />  
    <pagerstyle BackColor="white" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Size="smaller" />  
    <columns>
    <asp:HyperLinkColumn
        DataNavigateUrlField="t_id" DataNavigateUrlFormatString="s01insp_total_detail.aspx?t_id={0}"
        DataTextField="t_name" 
        Visible="True" target="total_detail" 
        headertext="名稱"/>    
    <asp:BoundColumn DataField="t_address" 
        HeaderText="地址" 
        ReadOnly="true" 
        Visible="True"/>    
    <asp:BoundColumn DataField="t_per" 
        HeaderText="人數" 
        ReadOnly="true" 
        Visible="True"/>    
    <asp:BoundColumn DataField="t_g_name" 
        HeaderText="行業" 
        ReadOnly="true" 
        Visible="True"/>    
    <asp:BoundColumn DataField="t_keyin" 
        HeaderText="更新者" 
        ReadOnly="true" 
        Visible="True"/>    
<asp:TemplateColumn HeaderText="時間" 
        Visible="True">
  <itemtemplate><%# FormatDateTime(DataSet1.FieldValue("t_update_time", Container), DateFormat.ShortDate) %></itemtemplate>
</asp:TemplateColumn>
    </columns>
  </asp:DataGrid>
</form>
<%
  session("cancel_insp") = ""
  session("cancel_insp") = request.url.tostring()
%>

</body>
</html>