<%@ Page Language="VB" ContentType="text/html" ResponseEncoding="utf-8" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<%@ Register TagPrefix="MM" Namespace="DreamweaverCtrls" Assembly="DreamweaverCtrls,version=1.0.0.0,publicKeyToken=836f606ede05d46a,culture=neutral" %>
<MM:DataSet 
id="DataSet1"
runat="Server"
IsStoredProcedure="false"
ConnectionString='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_STRING_s01") %>'
DatabaseType='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_DATABASETYPE_s01") %>'
CommandText='<%# "SELECT * FROM kc_calen_appen WHERE m_appen_id = ?" %>'
Debug="true"
>
  <Parameters>
    <Parameter  Name="@m_appen_id"  Value='<%# IIf((Request.QueryString("m_appen_id") <> Nothing), Request.QueryString("m_appen_id"), "") %>'  Type="Integer"   />
  </Parameters>
</MM:DataSet>
<MM:PageBind runat="server" PostBackBind="false" />
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.OleDb" %>
<%@ Import Namespace="System.IO" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<Html>
<style type="text/css">
<!--
.style1 {color: #FF0000}
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
      <td width="78%" bgcolor="#33FFFF"><Input name="file" Type="file" id="fileup" width="50" runat="server">
        <asp:TextBox Columns="20" ID="m_appen" ReadOnly="true" runat="server" text='<%# DataSet1.FieldValue("m_appen", Container) %>' />
        <asp:TextBox BackColor="#33FFFF" BorderWidth="0" Columns="3" ForeColor="#33FFFF" ID="m_num" ReadOnly="true" runat="server" Text='<%# DataSet1.FieldValue("m_num", Container) %>' />
		<asp:TextBox BackColor="#33FFFF" BorderWidth="0" Columns="3" ForeColor="#33FFFF" ID="m_appen_id" ReadOnly="true" runat="server" Text='<%# DataSet1.FieldValue("m_appen_id", Container) %>' />
	  </td>
    </tr>
    <tr>
      <td bgcolor="#FFFF33">檔案說明：</td>
      <td bgcolor="#33FFFF"><asp:TextBox Columns="50" ID="m_appen_con" Rows="2" runat="server" text='<%# DataSet1.FieldValue("m_appen_con", Container) %>' TextMode="MultiLine" /></td>
    </tr>
  </table>
  <br/>
    <asp:Button runat="server" Text="修改" OnClick="updatedata" />
    <asp:Button ID="Button1" runat="server" Text="取消" OnClick="goback" />    
<hr>
<p>
  <asp:Label ForeColor="#FF0000" id="Msg" runat="server" /></p>
<p>
  <label></label>
</p>
<p>
  <label></label>
</p>
</Form>

<script language="VB" runat="server">
    dim url
	Sub Page_Load(Src As Object, E As EventArgs)

	End Sub
 Sub goback(sender As Object, e As EventArgs)
  dim url_goback
   url_goback = "kc_cal_upfile.aspx?m_num=" & m_num.text 
   response.Redirect(url_goback)
  end sub 

      Dim SaveFileName as string

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
  Sub updateData(sender As Object, e As EventArgs) 
      Dim file As HttpPostedFile = FileUp.PostedFile
	  Dim MaxFileSize as Integer = 1500
      Dim ServerPath as String = Server.MapPath("s01data/")
      Dim FileSplit() As String = Split( File.FileName, "\" )
      Dim FileName As String = FileSplit(FileSplit.Length-1)

		if file.ContentLength <> 0 then
		If (file.ContentLength/1024) > MaxFileSize then
			Msg.Text ="上傳失敗" & "<br>" & "上傳檔案不能大於：" & MaxFileSize & "KB" & "<br>" & "上傳檔名長度不能多於25個字"
		exit sub
		End If
		if len(filename) > 25 then 
			Msg.Text ="上傳失敗" & "<br>" & "上傳檔案不能大於：" & MaxFileSize & "KB" & "<br>" & "上傳檔名長度不能多於25個字"
		exit sub
		 end if
      'if file.contenttype  = "image/pjpeg" or file.contenttype  = "image/gif" then
      'If file.ContentLength <> 0 Then
         SaveFileName = GetFileName(ServerPath,FileName)
		 File.SaveAs( ServerPath & SaveFileName )
		 m_appen.text = SaveFileName
      'End If
      Dim Conn As OleDbConnection
      Dim Cmd  As OleDbCommand

      Dim Provider = "Provider=Microsoft.Jet.OLEDB.4.0"
      Dim Database = "Data Source=" & Server.MapPath( "result/result.mdb" )
      Conn = New OleDbConnection( Provider & ";" & DataBase )
      Conn.Open()

      Dim SQL As String
      SQL = "UPDATE kc_calen_appen SET m_num=@m_num, m_appen=@m_appen, m_appen_con= @m_appen_con where m_appen_id=@m_appen_id"
      Cmd = New OleDbCommand( SQL, Conn )

      Cmd.Parameters.Add( New OleDbParameter("@m_num", OleDbType.integer))
      Cmd.Parameters.Add( New OleDbParameter("@m_appen", OleDbType.Char, 30))
      Cmd.Parameters.Add( New OleDbParameter("@m_appen_con", OleDbType.Char, 60))
      Cmd.Parameters.Add( New OleDbParameter("@m_appen_id", OleDbType.integer))

      Cmd.Parameters("@m_num").value = val(m_num.text)
      Cmd.Parameters("@m_appen").value = m_appen.text
      Cmd.Parameters("@m_appen_con").value = m_appen_con.text
      Cmd.Parameters("@m_appen_id").value = val(m_appen_id.text)

      Cmd.ExecuteNonQuery()
      If Err.Number <> 0 Then
         Msg.Text = Err.Description
      Else
		 url = "kc_cal_upfile.aspx?m_num=" & m_num.text
      End If

      Conn.Close()
      'Else
         Msg.Text = "<font color='red'>上傳失敗，請重新上傳，檔案大小不能超過1.5M</font>"
      'end if
    else
      Dim Conn As OleDbConnection
      Dim Cmd  As OleDbCommand

      Dim Provider = "Provider=Microsoft.Jet.OLEDB.4.0"
      Dim Database = "Data Source=" & Server.MapPath( "result/result.mdb" )
      Conn = New OleDbConnection( Provider & ";" & DataBase )
      Conn.Open()

      Dim SQL As String
      SQL = "UPDATE kc_calen_appen SET m_num=@m_num, m_appen=@m_appen, m_appen_con= @m_appen_con where m_appen_id=@m_appen_id"
      Cmd = New OleDbCommand( SQL, Conn )

      Cmd.Parameters.Add( New OleDbParameter("@m_num", OleDbType.integer))
      Cmd.Parameters.Add( New OleDbParameter("@m_appen", OleDbType.Char, 30))
      Cmd.Parameters.Add( New OleDbParameter("@m_appen_con", OleDbType.Char, 60))
      Cmd.Parameters.Add( New OleDbParameter("@m_appen_id", OleDbType.integer))

      Cmd.Parameters("@m_num").value = val(m_num.text)
      Cmd.Parameters("@m_appen").value = m_appen.text
      Cmd.Parameters("@m_appen_con").value = m_appen_con.text
      Cmd.Parameters("@m_appen_id").value = val(m_appen_id.text)

      Cmd.ExecuteNonQuery()
      If Err.Number <> 0 Then
         Msg.Text = Err.Description
      Else
		 url = "kc_cal_upfile.aspx?m_num=" & m_num.text
      End If

      Conn.Close()
	  end if	   
   End Sub
</script>
<%
if url <> "" then
%>
<script language = "JavaScript">
<!--
//將已經上傳的檔案名稱傳回到新增照片頁面的Photo_Picture
//隱藏欄位，用於交檔案名稱儲存到資料庫中。
location.href="<% =url %>";
//-->
</Script>
<%
End If
%>
</Body>
</Html>
