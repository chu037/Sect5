<%@ Page Language="VB" ContentType="text/html" ResponseEncoding="utf-8" %>
<%@ Register TagPrefix="MM" Namespace="DreamweaverCtrls" Assembly="DreamweaverCtrls,version=1.0.0.0,publicKeyToken=836f606ede05d46a,culture=neutral" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.OleDb" %>
<%@ Import Namespace="System.Xml" %>

<MM:DataSet 
id="DataSet1"
runat="Server"
IsStoredProcedure="false"
ConnectionString='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_STRING_s01") %>'
DatabaseType='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_DATABASETYPE_s01") %>'
CommandText='<%# "SELECT * FROM s01_total WHERE t_per between ? and ? and t_g_1 like ? and t_g_2 like ? and t_g_3 like ? and t_group.t_g_num like ? and t_address like ? and t_name like ? and t_pre like ? ORDER BY t_per DESC" %>'
Debug="true" PageSize="50"
>
  <Parameters>

    <Parameter  Name="@t_person1"  Value='<%# IIf((Request.QueryString("t_person1") <> Nothing), Request.QueryString("t_person1"), "0") %>'  Type="integer"   />
    <Parameter  Name="@t_person2"  Value='<%# IIf((Request.QueryString("t_person2") <> Nothing), Request.QueryString("t_person2"), "99999") %>'  Type="integer"   />
<Parameter  Name="@t_g_1"  Value='<%# IIf((Request.QueryString("t_g_1") <> Nothing), Request.QueryString("t_g_1"), "%") %>'  Type="WChar"   />
<Parameter  Name="@t_g_2"  Value='<%# IIf((Request.QueryString("t_g_2") <> Nothing), Request.QueryString("t_g_2"), "%") %>'  Type="WChar"   />
<Parameter  Name="@t_g_3"  Value='<%# IIf((Request.QueryString("t_g_3") <> Nothing), Request.QueryString("t_g_3"), "%") %>'  Type="WChar"   />
<Parameter  Name="@t_g_num"  Value='<%# IIf((Request.QueryString("t_g_num") <> Nothing), Request.QueryString("t_g_num"), "%") %>'  Type="WChar"   />
<Parameter  Name="@t_address"  Value='<%# "%" + (IIf((Request.QueryString("t_address") <> Nothing), Request.QueryString("t_address"), "%")) + "%" %>'  Type="WChar"   />
<Parameter  Name="@t_name"  Value='<%# "%" + (IIf((Request.QueryString("t_name") <> Nothing), Request.QueryString("t_name"), "%")) + "%" %>'  Type="WChar"   />
<Parameter  Name="@t_pre"  Value='<%# "%" + (IIf((Request.QueryString("t_pre") <> Nothing), Request.QueryString("t_pre"), "%")) + "%" %>'  Type="WChar"   />

</Parameters>
</MM:DataSet>
<MM:DataSet 
id="DataSet2"
runat="Server"
IsStoredProcedure="false"
ConnectionString='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_STRING_s01") %>'
DatabaseType='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_DATABASETYPE_s01") %>'
CommandText='<%# "SELECT *  FROM s01_total  WHERE t_per between ? and ? and t_g_1 like ? and t_g_2 like ? and t_g_3 like ? and t_group.t_g_num like ? and t_address like ? and t_name like ? and t_pre like ? ORDER BY t_per DESC" %>'
Debug="true"
>
  <Parameters>
    <Parameter  Name="@t_person1"  Value='<%# IIf((Request.QueryString("t_person1") <> Nothing), Request.QueryString("t_person1"), "0") %>'  Type="integer"   />
    <Parameter  Name="@t_person2"  Value='<%# IIf((Request.QueryString("t_person2") <> Nothing), Request.QueryString("t_person2"), "99999") %>'  Type="integer"   />
<Parameter  Name="@t_g_1"  Value='<%# IIf((Request.QueryString("t_g_1") <> Nothing), Request.QueryString("t_g_1"), "%") %>'  Type="WChar"   />
<Parameter  Name="@t_g_2"  Value='<%# IIf((Request.QueryString("t_g_2") <> Nothing), Request.QueryString("t_g_2"), "%") %>'  Type="WChar"   />
<Parameter  Name="@t_g_3"  Value='<%# IIf((Request.QueryString("t_g_3") <> Nothing), Request.QueryString("t_g_3"), "%") %>'  Type="WChar"   />
<Parameter  Name="@t_g_num"  Value='<%# IIf((Request.QueryString("t_g_num") <> Nothing), Request.QueryString("t_g_num"), "%") %>'  Type="WChar"   />
<Parameter  Name="@t_address"  Value='<%# "%" + (IIf((Request.QueryString("t_address") <> Nothing), Request.QueryString("t_address"), "%")) + "%" %>'  Type="WChar"   />
<Parameter  Name="@t_name"  Value='<%# "%" + (IIf((Request.QueryString("t_name") <> Nothing), Request.QueryString("t_name"), "%")) + "%" %>'  Type="WChar"   />
<Parameter  Name="@t_pre"  Value='<%# "%" + (IIf((Request.QueryString("t_pre") <> Nothing), Request.QueryString("t_pre"), "%")) + "%" %>'  Type="WChar"   /></Parameters></MM:DataSet>
<MM:DataSet 
id="DataSet3"
runat="Server"
IsStoredProcedure="false"
ConnectionString='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_STRING_s01") %>'
DatabaseType='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_DATABASETYPE_s01") %>'
CommandText='<%# "SELECT * FROM s01cs" %>'
Debug="true"
></MM:DataSet>
<MM:PageBind runat="server" PostBackBind="false" />

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<script runat="server">
		Sub Page_Load(Src As Object, E As EventArgs)
		If Session("MM_username")="" then
			Response.Write("您還沒有登入呢！，請點擊")
			Response.Write("<a href=s01login.aspx>登入</a>")
			Response.End()
		End If
		if ispostback then
		exit sub
		end if
	  total_cancel.text = session("total_cancel")
	  session("total_cancel") = ""
	End Sub

   Sub sortagain(sender As Object, e As DataGridSortCommandEventArgs)
      If SortField.Text = e.SortExpression Then
         If SortType.Text = "" Then
            SortType.Text = " Desc"
         Else
            SortType.Text = ""
         End If
      Else
         SortField.Text = e.SortExpression
         SortType.Text  = ""
      End If
   End Sub

   Sub starsearch(sender As Object, e As EventArgs) 
	  Response.Redirect( total_cancel.text ) ' 使用Server.Transfer亦可
   End Sub

Protected Sub Button3_Click(ByVal sender As Object, ByVal e As System.EventArgs) 
 response.Redirect("s01insp_plant_search_index.aspx")
 end sub

Protected Sub Button2_Click(ByVal sender As Object, ByVal e As System.EventArgs) 

        Dim ExportWriter As New IO.StringWriter()
        Dim ExportHtmlTextWriter As New HtmlTextWriter(ExportWriter)
        'Dim xlsApp As Object  
        'Dim strValue As String
		'Dim doc as new Xmldocument()  
		'DataGrid1.AllowPaging = False
        'DataGrid1.AllowSorting = False
        DataGrid2.visible = true
		DisableControls(DataGrid2)
        DataGrid2.RenderControl(ExportHtmlTextWriter)

        Response.AppendHeader("Content-Disposition", "attachment;filename=" + Date.Today.ToString("yyyy-MM-dd") + ".xls")
 'xlsApp = CreateObject("Excel.Application")   
 'xlsApp.Workbooks.Open("C:\Book1.xls")   

        'Response.Charset = System.Text.Encoding.GetEncoding("UTF-8")
       Response.Charset = "UTF-8"
       'Response.HeaderEncoding= "utf-8"
       Response.ContentType = "application/vnd.ms-excel"
        'Response.ContentType = "text/xml"
        'Response.ClearHeaders()
		Response.Write(ExportWriter.ToString())
		'Response.Write(ExportWriter)
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

function cs_chk(address)
dim cs_name() = {"新興區","前金區","苓雅區","鹽埕區","鼓山區","旗津區","前鎮區","三民區","楠梓區","小港區","左營區","仁武區","大社區","岡山區","路竹區","阿蓮區","田寮區","燕巢區","橋頭區","梓官區","彌陀區","永安區","湖內區","鳳山區","大寮區","林園區","鳥松區","大樹區","旗山區","美濃區","六龜區","內門區","杉林區","甲仙區","桃源區","那瑪夏區","茂林區","茄萣區"}
dim cs_num() = {"800","801","802","803","804","805","806","807","811","812","813","814","815","820","821","822","823","824","825","826","827","828","829","830","831","832","833","840","842","843","844","845","846","847","848","849","851","852"}
if address <> "" then
dim i
for i = 0 to ubound(cs_name) 
if instr(address, cs_name(i)) > 0 then
cs_chk = cs_num(i)
exit for
exit function
end if
next
end if
end function

</script>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<title>事業單位查詢結果</title><!--參數應照順序 %為萬用字元 like 前後 + "%"-->
<style type="text/css">
<!--
.style2 {
	color: #CC6600;
	font-weight: bold;
}
-->
</style>
</head>
<body>
<form runat="server">
  <p>共找到 <span class="style2"><%= DataSet1.RecordCount %></span>筆資料
<asp:Button ID="Button1" runat="server" Text="回檢查名冊查詢" OnClick="starsearch" />
  <a href="s01insp_logout.aspx" target="_self">
  <asp:Button ID="Button2" runat="server" Text="將資料彙出至EXCEL" OnClick="Button2_Click" />  
 登出</a>
  <asp:TextBox Columns="360" ID="total_cancel" runat="server" Visible="false" />
</p>
  <p>
    <asp:DataGrid AllowCustomPaging="true" 
  AllowPaging="true" 
  AllowSorting="true" 
  AutoGenerateColumns="false" 
  CellPadding="3" 
  CellSpacing="0" 
  DataSource="<%# DataSet1.DefaultView %>" Enabled="true" id="DataGrid1" PagerStyle-Mode="NumericPages" PageSize="<%# DataSet1.PageSize %>" 
  runat="server" 
  ShowFooter="false" 
  ShowHeader="true" Width="100%" OnPageIndexChanged="DataSet1.OnDataGridPageIndexChanged" virtualitemcount="<%# DataSet1.RecordCount %>" 
>
      <HeaderStyle HorizontalAlign="left" BackColor="#E8EBFD" ForeColor="#3D3DB6" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Bold="true" Font-Size="smaller" />      
<ItemStyle BackColor="#F2F2F2" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Size="smaller" />      
<AlternatingItemStyle BackColor="#E5E5E5" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Size="smaller" />      
<FooterStyle HorizontalAlign="left" BackColor="#E8EBFD" ForeColor="#3D3DB6" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Bold="true" Font-Size="smaller" />      
<PagerStyle BackColor="white" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Size="smaller" />      
<Columns>
      <asp:HyperLinkColumn DataNavigateUrlField="t_id" DataNavigateUrlFormatString="s01insp_total_detail.aspx?t_id={0}"
        DataTextField="t_name" 
        Visible="True" target="total_detail" 
        HeaderText="事業單位" 
		/>      
      <asp:BoundColumn DataField="t_address" 
        HeaderText="地址" 
        ReadOnly="true" 
        Visible="True"/>      
      <asp:BoundColumn DataField="t_address_m" 
        HeaderText="通訊地址" 
        ReadOnly="true" 
        Visible="True"/>      
      <asp:BoundColumn DataField="t_tel" 
        HeaderText="電話" 
        ReadOnly="true" 
        Visible="True"/>      
      <asp:BoundColumn DataField="t_per" 
        HeaderText="人數" 
        ReadOnly="true" 
        Visible="True"/>      
      <asp:BoundColumn DataField="t_group.t_g_name" 
        HeaderText="行業別" 
        ReadOnly="true" 
        Visible="True"/>      
</Columns>
    </asp:DataGrid>
	<asp:DataGrid 
  AllowPaging="false" 
  AllowSorting="false" 
  AutoGenerateColumns="false" 
  CellPadding="3" 
  CellSpacing="0" 
  DataSource="<%# DataSet2.DefaultView %>" Enabled="true" id="DataGrid2" 
  runat="server" 
  ShowFooter="false" 
  ShowHeader="true" Visible="false" Width="100%" 
>
      <HeaderStyle HorizontalAlign="left" BackColor="#E8EBFD" ForeColor="#3D3DB6" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Bold="true" Font-Size="smaller" />      
<ItemStyle BackColor="#F2F2F2" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Size="smaller" />      
<AlternatingItemStyle BackColor="#E5E5E5" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Size="smaller" />      
<FooterStyle HorizontalAlign="left" BackColor="#E8EBFD" ForeColor="#3D3DB6" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Bold="true" Font-Size="smaller" />      
<PagerStyle BackColor="white" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Size="smaller" />      
<Columns>
<asp:BoundColumn DataField="t_name" 
        HeaderText="事業單位" 
        ReadOnly="true" 
        Visible="True"/>
<asp:TemplateColumn HeaderText="區號" 
        Visible="True">
  <ItemTemplate><%# cs_chk(DataSet2.FieldValue("t_address", Container)) %></ItemTemplate>
</asp:TemplateColumn>
<asp:BoundColumn DataField="t_address" 
        HeaderText="地址" 
        ReadOnly="true" 
        Visible="True"/>
<asp:TemplateColumn HeaderText="通訊區號" 
        Visible="True">
  <ItemTemplate><%# cs_chk(DataSet2.FieldValue("t_address_m", Container)) %></ItemTemplate>
</asp:TemplateColumn>
<asp:BoundColumn DataField="t_address_m" 
       HeaderText="通訊地址" 
        ReadOnly="true" 
        Visible="True"/>      
<asp:BoundColumn DataField="t_tel" 
        HeaderText="電話" 
        ReadOnly="true" 
        Visible="True"/>
<asp:BoundColumn DataField="t_per" 
        HeaderText="人數" 
        ReadOnly="true" 
        Visible="True"/>
<asp:BoundColumn DataField="t_group.t_g_name" 
        HeaderText="行業別" 
        ReadOnly="true" 
        Visible="True"/>
</Columns>
    </asp:DataGrid>
    <asp:Label runat="server" id="SortField" Text="專案" Visible="False" />    
    <asp:Label runat="server" id="SortType" Text="" Visible="False" />    
</p>
</form>
</body>
</html>

