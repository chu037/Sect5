<%@ Page Language="VB" ContentType="text/html" ResponseEncoding="utf-8" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<%@ Register TagPrefix="MM" Namespace="DreamweaverCtrls" Assembly="DreamweaverCtrls,version=1.0.0.0,publicKeyToken=836f606ede05d46a,culture=neutral" %>
<MM:DataSet 
id="DataSet2"
runat="Server"
IsStoredProcedure="false"
ConnectionString='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_STRING_s01") %>'
DatabaseType='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_DATABASETYPE_s01") %>'
CommandText='<%# "SELECT * FROM calen_appen WHERE m_num = ? ORDER BY m_appen_id DESC" %>'
Debug="true"
><Parameters>
<Parameter  Name="@m_num"  Value='<%# IIf((Request.QueryString("m_num") <> Nothing), Request.QueryString("m_num"), "") %>'  Type="Integer"   /></Parameters></MM:DataSet>
<MM:PageBind runat="server" PostBackBind="true" />
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.OleDb" %>
<%@ Import Namespace="System.IO" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<Html>
<style type="text/css">
<!--
.style1 {color: #FF0000}
.style2 {
	color: #FF6600;
	font-size: 14px;
}
-->
</style>
<Body BgColor=White>
<H3>檔案上傳<span class="style1">檔案大小勿超過1.5M</span>
  <Hr></H3>

<Form name='form2' method='POST' Enctype="multipart/form-data" runat="server">
  <p>
    <label></label>
  </p>
  <table width="100%" border="1" cellspacing="0" cellpadding="0">
    <tr>
      <td width="22%" bgcolor="#FFFF33">請輸入完整檔案路徑：</td>
      <td width="78%" bgcolor="#33FFFF"><input name="file" type="file" id="fileup" width="50" runat="server">
      <asp:TextBox ID="m_appen" runat="server" AutoPostBack="true" ReadOnly="true" />
        
        <asp:TextBox BackColor="#33FFFF" BorderWidth="0" Columns="6" ForeColor="#33FFFF" ID="m_num" ReadOnly="true" runat="server" /></td>
    </tr>
    <tr>
      <td bgcolor="#FFFF33">檔案說明：</td>
      <td bgcolor="#33FFFF"><asp:TextBox Columns="50" ID="m_appen_con" Rows="2" runat="server" TextMode="MultiLine" /></td>
    </tr>
  </table>
  <br/>
    <asp:Button ID="Button5" runat="server" Text="上傳檔案" OnClick="insertdata" />      
    <asp:Button ID="Button2" runat="server" Text="回修改行程" OnClick="goback" />    
<hr>
<p>
  <asp:Label ForeColor="#FF0000" id="Msg" runat="server" />
  <br>
  <asp:Button ID="Button1" runat="server" Text="確定" OnClick="confirm" />  
</p>
<strong><span class="style2">相關檔案如下:</span><br>
</strong>
<asp:DataGrid 
  AllowPaging="false" 
  AllowSorting="False" 
  AutoGenerateColumns="false" 
  CellPadding="3" 
  CellSpacing="0" 
  DataSource="<%# DataSet2.DefaultView %>" id="DataGrid1" 
  runat="server" 
  ShowFooter="false" 
  ShowHeader="true" Width="100%" 
>
      <HeaderStyle HorizontalAlign="center" BackColor="#E8EBFD" ForeColor="#3D3DB6" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Bold="true" Font-Size="smaller" />    
      <ItemStyle BackColor="#F2F2F2" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Size="smaller" />    
      <AlternatingItemStyle BackColor="#E5E5E5" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Size="smaller" />    
      <FooterStyle HorizontalAlign="center" BackColor="#E8EBFD" ForeColor="#3D3DB6" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Bold="true" Font-Size="smaller" />    
      <PagerStyle BackColor="white" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Size="smaller" />    
      <Columns>
      <asp:HyperLinkColumn
        DataNavigateUrlField="m_appen"
        DataNavigateUrlFormatString="/web1/sec05/s01data/{0}"
        DataTextField="m_appen" 
        Visible="True" target="show_photo" 
        HeaderText="檔名"/>      
<asp:BoundColumn DataField="m_appen_con" 
        HeaderText="檔案說明" 
        ReadOnly="true" 
        Visible="True"/>      
<asp:TemplateColumn HeaderText="修改" 
        Visible="True">
  <ItemTemplate>
    <input name="Submit3" type="submit" onclick="MM_goToURL('self','s01cal_upfile_update.aspx?m_appen_id=<%# DataSet2.FieldValue("m_appen_id", Container) %>');return document.MM_returnValue" value="修改" />
  </ItemTemplate>
</asp:TemplateColumn>
      </Columns>
  </asp:DataGrid>
</p>
<p>
  <label></label>
</p>
<p>
  <label></label>
</p>
</Form>

<script language="VB" runat="server">
    dim url, m_numcheck
	Sub Page_Load(Src As Object, E As EventArgs)
	m_numcheck = request("m_num")
	 if m_numcheck = "" then
      response.Redirect("calendartest02.aspx")
	 else
	 m_num.text = m_numcheck
	end if
	if not ispostback then
	 button1.enabled = false
	 button1.visible = false
	end if 
	End Sub
 Sub confirm(Src As Object, E As EventArgs)
  msg.text = ""
  button1.enabled = false
  button1.visible = false
  button2.enabled = true
  button2.visible = true
  button5.enabled = true
  button5.visible = true

  end sub
 Sub goback(sender As Object, e As EventArgs)
  dim url_goback
  if m_numcheck <> nothing then
   url_goback = "s01cal_update.aspx?m_num=" & m_numcheck 
   response.Redirect(url_goback)
  end if
  end sub 

      Dim SaveFileName as string
   Sub UploadFile(sender As Object, e As EventArgs)
      Dim file As HttpPostedFile = FileUp.PostedFile
	  Dim MaxFileSize as Integer = 3000
      Dim ServerPath as String = Server.MapPath("s01data/")
		
		If (file.ContentLength/1024) > MaxFileSize then
			Msg.Text ="上傳檔案不能大於：" & MaxFileSize & "KB"
			
		End If
      'if file.contenttype  = "image/pjpeg" or file.contenttype  = "image/gif" then
      If file.ContentLength <> 0 Then
         Msg.Text  = "<font color='#336633'>檔案上傳成功</font>"
		 Msg.Text &= "<br>大小: " & File.ContentLength
         Msg.Text &= "<br>類型: " & File.ContentType
         Msg.Text &= "<br>名稱: " & File.FileName

         Dim FileSplit() As String = Split( File.FileName, "\" )
         Dim FileName As String = FileSplit(FileSplit.Length-1)
         SaveFileName = GetFileName(ServerPath,FileName)
		 File.SaveAs( ServerPath & SaveFileName )
		 m_appen.text = SaveFileName
      Else
         Msg.Text = "<font color='red'>上傳失敗，請重新上傳，且檔案大小不能超過1.5M</font>"
      End If
      'end if
  End Sub

Function GetFileName(ServerPath, FileName)
        Dim leftFileName = ""
        Dim rightFileName = ""
		Dim i
        For i = Len(FileName) To 1 Step -1
            If Mid(FileName, i, 1) = Chr(Asc(".")) Then
                leftFileName = Left(FileName, i-1)
                rightFileName = Right(FileName, Len(FileName)-i+1)
                Exit For
            End If
        Next
        For i = 0 to 9999 
             if file.Exists(ServerPath & leftFileName & "_" & i & rightFileName ) = False then
                 FileName = leftFileName  & "_" & i & rightFileName
                 Exit For
             End If
        Next
		GetFileName = FileName
	End Function
  Sub InsertData(sender As Object, e As EventArgs) 
      Dim file As HttpPostedFile = FileUp.PostedFile
	  Dim MaxFileSize as Integer = 1500
      Dim ServerPath as String = Server.MapPath("s01data/")
      Dim FileSplit() As String = Split( File.FileName, "\" )
      Dim FileName As String = FileSplit(FileSplit.Length-1)

		If (file.ContentLength/1024) > MaxFileSize then
			Msg.Text ="上傳失敗" & "<br>" & "上傳檔案不能大於：" & MaxFileSize & "KB" & "<br>" & "上傳檔名長度不能多於25個字"
		exit sub
		End If
		if len(filename) > 25 then 
			Msg.Text ="上傳失敗" & "<br>" & "上傳檔案不能大於：" & MaxFileSize & "KB" & "<br>" & "上傳檔名長度不能多於25個字"
		exit sub
		 end if	

      'if file.contenttype  = "image/pjpeg" or file.contenttype  = "image/gif" then
      If file.ContentLength <> 0 Then
         Msg.Text  = "<font color='#336633'>檔案上傳成功</font>"
		 Msg.Text &= "<br>大小: " & File.ContentLength
         Msg.Text &= "<br>類型: " & File.ContentType
         Msg.Text &= "<br>名稱: " & File.FileName

         SaveFileName = GetFileName(ServerPath,FileName)
		 File.SaveAs( ServerPath & SaveFileName )
		 m_appen.text = SaveFileName
      End If
      'Else
         'Msg.Text = "<font color='red'>上傳失敗，請重新上傳，只能上傳gif或jpg檔，且檔案大小不能超過1.5M</font>"
		'exit sub 
      'end if
	if SaveFileName <> nothing then
	if IsValid Then

      Dim Conn As OleDbConnection
      Dim Cmd  As OleDbCommand

      Dim Provider = "Provider=Microsoft.Jet.OLEDB.4.0"
      Dim Database = "Data Source=" & Server.MapPath( "result/result.mdb" )
      Conn = New OleDbConnection( Provider & ";" & DataBase )
      Conn.Open()

      Dim SQL As String
      SQL = "Insert Into calen_appen (m_num, m_appen, m_appen_con) Values(@m_num, @m_appen, @m_appen_con)"
      Cmd = New OleDbCommand( SQL, Conn )

      Cmd.Parameters.Add( New OleDbParameter("@m_num", OleDbType.integer))
      Cmd.Parameters.Add( New OleDbParameter("@m_appen", OleDbType.Char, 30))
      Cmd.Parameters.Add( New OleDbParameter("@m_appen_con", OleDbType.Char, 60))

      Cmd.Parameters("@m_num").value = val(m_num.text)
      Cmd.Parameters("@m_appen").value = m_appen.text
      Cmd.Parameters("@m_appen_con").value = m_appen_con.text

      Cmd.ExecuteNonQuery()
      If Err.Number <> 0 Then
         Msg.Text = Err.Description
      Else
		 button1.enabled = true
		 button1.visible = true
		 m_appen.text = ""
		 m_appen_con.text = ""
		 button2.enabled = false
		 button2.visible = false
		 button5.enabled = false
		 button5.visible = false
      
	  End If

      Conn.Close()
      End If
	  else
	  msg.text = "尚未點選上傳檔案"
	  end if
   End Sub
</script>
<script language = "JavaScript">
<!--
function w_back()
{
var x1;
x1 == document.form_update.m_date.value;
location.href = "s01cal_selday.aspx?m_date=" & x1;
}

function Mcheck(){
	if (document.form_update.m_note.value=="") {
        window.alert("請輸入內容");
        return false }
    if (document.form_update.m_hours.value=="") {
        window.alert("請輸入預定時數");
        return false }
	if (isNaN(document.form_update.m_hours.value)) {
        window.alert("時數請輸入數值");
        return false }
	 return true;
}

function MM_goToURL() { //v3.0
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"'");
}

function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
}
//-->
</Script>
</Body>
</Html>