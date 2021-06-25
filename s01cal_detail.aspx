<%@ Page Language="VB" ContentType="text/html" ResponseEncoding="utf-8" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<%@ Register TagPrefix="MM" Namespace="DreamweaverCtrls" Assembly="DreamweaverCtrls,version=1.0.0.0,publicKeyToken=836f606ede05d46a,culture=neutral" %>
<MM:DataSet 
id="DataSet1"
runat="Server"
IsStoredProcedure="false"
ConnectionString='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_STRING_s01") %>'
DatabaseType='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_DATABASETYPE_s01") %>'
CommandText='<%# "SELECT m_date, m_date_en, m_detail, m_hours, m_messeng, m_note, m_num, m_time, m_time_en FROM s01_calen WHERE m_num = ?" %>'
Debug="true"
><Parameters>
<Parameter  Name="@m_num"  Value='<%# IIf((Request.QueryString("m_num") <> Nothing), Request.QueryString("m_num"), "") %>'  Type="Integer"   /></Parameters></MM:DataSet>
<MM:DataSet 
id="DataSet2"
runat="Server"
IsStoredProcedure="false"
ConnectionString='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_STRING_s01") %>'
DatabaseType='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_DATABASETYPE_s01") %>'
CommandText='<%# "SELECT * FROM calen_appen WHERE m_num = ? ORDER BY m_appen_id ASC" %>'
Debug="true"
>
  <Parameters>
    <Parameter  Name="@m_num"  Value='<%# IIf((Request.QueryString("m_num") <> Nothing), Request.QueryString("m_num"), "") %>'  Type="Integer"   />
  </Parameters>
</MM:DataSet>
<MM:PageBind runat="server" PostBackBind="true" />
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8 />
<title>單一行程</title>
<script language="VB" runat="server">
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
 function showweek(t)
 if t <> ""
 showweek = weekdayname(datepart("w",t))
 end if
 end function

 function showmessage(vt2,vt3)
 if vt2 <> ""
 showmessage = vt2 & "<br/>" & vt3 & "小時"
 end if
 end function
  Function Clean(str)
   str=Replace(str, vbCrLf, "<br>")
   Clean=Replace(str, chr(32), "&nbsp;&nbsp;")
  End Function

</script>
<script language = "JavaScript">
<!--



function w_back()
{
location.href = "calendartest02.aspx"
}
//-->
</Script>
<style type="text/css">
<!--
.style1 {color: #3300FF}
.style2 {color: #CC3300}
-->
</style>
</head>
<body>
<table width="655" border="1" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF" id="t00">
  <tr>
    <td width="13%" bgcolor="#FFFFFF"><div align="center">日期</div></td>
    <td width="87%" bgcolor="#FFFFFF"><%# showdate(DataSet1.FieldValue("m_date", Container)) %><span class="style1"><%# showweek(DataSet1.FieldValue("m_date", Container)) %></span><%# showtime(DataSet1.FieldValue("m_time", Container)) %>~<br /><%# showdate(DataSet1.FieldValue("m_date_en", Container)) %><span class="style1"><%# showweek(DataSet1.FieldValue("m_date_en", Container)) %></span><%# showtime(DataSet1.FieldValue("m_time_en", Container)) %>，時數:<span class="style2"><%# DataSet1.FieldValue("m_hours", Container) %></span>小時</td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF"><div align="center">行程</div></td>
    <td bgcolor="#FFFFFF"><%# DataSet1.FieldValue("m_messeng", Container) %></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF"><div align="center">主題</div></td>
    <td bgcolor="#FFFFFF"><%# DataSet1.FieldValue("m_note", Container) %></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF"><div align="center">詳細內容</div></td>
    <td bgcolor="#FFFFFF"><%# clean(DataSet1.FieldValue("m_detail", Container)) %></td>
  </tr>
  <tr bgcolor="#FFFFFF">
    <td colspan="2" align="center">相關圖照</td>
  </tr>
</table>
 
  <form runat="server">
    <asp:DataList id="DataList1" 
runat="server" 
RepeatColumns="3" 
RepeatDirection="Horizontal" 
RepeatLayout="Table" 
DataSource="<%# DataSet2.DefaultView %>" >
      <ItemTemplate>
      <table width="217" border="0" cellpadding="0" cellspacing="0" id="t01">
        <tr>

          <td width="217" align="center" bgcolor="#FFFFFF" ><a href="s01data/<%# DataSet2.FieldValue("m_appen", Container) %>" target="win2" ><img src="s01data/<%# DataSet2.FieldValue("m_appen", Container) %>" alt="<%# DataSet2.FieldValue("m_appen", Container) %>" name="p01" width="217" height="173" border="0" align="baseline" id="p01" /></a></td>
        </tr>
      </table> </ItemTemplate>
    </asp:DataList>
  </form>

  <form id="form1" name="form1" method="post" action="">
<p>
  <input name="Submit" type="button" onclick="window.close()" value="關閉視窗" />
</p>
</form>
</body>
</html>