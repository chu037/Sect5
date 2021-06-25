<%@ Page Language="VB" ContentType="text/html" ResponseEncoding="big5" %>
<%@ Register TagPrefix="MM" Namespace="DreamweaverCtrls" Assembly="DreamweaverCtrls,version=1.0.0.0,publicKeyToken=836f606ede05d46a,culture=neutral" %>
<MM:DataSet 
id="DataSet1"
runat="Server"
IsStoredProcedure="false"
ConnectionString='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_STRING_s01") %>'
DatabaseType='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_DATABASETYPE_s01") %>'
CommandText='<%# "SELECT case_end, case_id, case_start FROM s01_case WHERE case_id = ?" %>'
Debug="true"
>
  <Parameters>
<Parameter  Name="@case_id"  Value='<%# IIf((Request.QueryString("case_id") <> Nothing), Request.QueryString("case_id"), "") %>'  Type="Integer"   /></Parameters></MM:DataSet>
<MM:PageBind runat="server" PostBackBind="false" />
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5" />
<title>選擇日期</title>
    <script language="VB" runat="server">
        dim x1, sel_s, sel_e, sel_s1, sel_e1
        Sub Page_Load(sender As Object, e As EventArgs)
		If Session("MM_chief_num")="" then
			Response.Write("您還沒有登入呢！，請點擊")
			Response.Write("<a href=s01dellogin.aspx>登入</a>")
			Response.End()
		End If

		  x1 = request ("s_date")
		   select case x1
		    case 1
			 sel_s = request("case_start")
			case 2
			 sel_e = request("case_end")
		   end select
		End Sub

		Sub Date_Selected(sender As Object, e As EventArgs)
          sel.Text = Calendar1.SelectedDate.ToShortDateString
        End Sub
     

		Sub d_t(sender As Object, e As EventArgs)
		   select case x1
		    case 1
			 sel_s1 = sel.text
			 sel_e1 = ""
			case 2
			 sel_e1 = sel.text
			 sel_s1 = ""
           end select
		End sub
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

    </script>

</head>

<body>

    <h3><font face="Verdana, 新細明體">請點選日期：</font></h3>

    <form runat="server">
      <p>
        <asp:Calendar id=Calendar1 runat="server"
            onselectionchanged="Date_Selected"
            Font-Name="Arial" Font-Size="12px"
            Height="180px" Width="200px"
            SelectorStyle-BackColor="gainsboro"
            TodayDayStyle-BackColor="#ffcc33"
            DayHeaderStyle-BackColor="#ffcccc"
            OtherMonthDayStyle-ForeColor="#ffffcc"
            TitleStyle-BackColor="#cccc66"
            TitleStyle-Font-Bold="True"
            TitleStyle-Font-Size="12px"
            SelectedDayStyle-BackColor="Navy"
            SelectedDayStyle-Font-Bold="True"
            />

        <br />
        日期：
        <asp:TextBox ID="sel" ReadOnly="true" runat="server"/>        
        <br />
        <asp:TextBox ID="case_start" ReadOnly="true" runat="server" text='<%# showdate(DataSet1.FieldValue("case_start", Container)) %>' Visible="false" />      
      <br />
      <asp:TextBox ID="case_end" ReadOnly="true" runat="server" Text='<%# showdate(DataSet1.FieldValue("case_end", Container)) %>' Visible="false" />      
      <br />

      <asp:Button ID="Button1" runat="server" Text="確認" OnClick="d_t" />      
<input name="Submit2" type="button" onclick="self.close()" value="取消" />
      <p>
      
</form>
<%
if sel_s1 <> ""
%>
<script language = "JavaScript">
<!--
opener.document.form_update.case_start.value ="<% =sel_s1 %>";
self.close();

//-->
</Script>

<%
end if		
%> 
<%
if sel_e1 <> ""
%>
<script language = "JavaScript">
<!--
opener.document.form_update.case_end.value ="<% =sel_e1 %>";
self.close();

//-->
</Script>

<%
end if		
%> 
</body>
</html>
