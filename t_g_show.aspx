<%@ Page Language="VB" ContentType="text/html" ResponseEncoding="utf-8" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.OleDb" %>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>無標題文件</title>
<script language="VB" runat="server">
  Sub page_load(sender As Object, e As EventArgs)
      Dim Conn As OleDbConnection
      Dim Cmd  As OleDbCommand
	  dim rd as OleDbDataReader

      Dim Provider = "Provider=Microsoft.Jet.OLEDB.4.0"
      Dim Database = "Data Source=" & Server.MapPath( "result/result.mdb" )
      Conn = New OleDbConnection( Provider & ";" & DataBase )
      Conn.Open()

      Dim SQL As String
	  SQL = "Select * From t_group"
      Cmd = New OleDbCommand( SQL, Conn )
	  
	  rd = cmd.ExecuteReader()
	  
	  call outtotb(rd)
	  conn.close()
	  end sub

  Sub outtotb(rd As OleDbDataReader)
   dim i as integer
   dim row as tablerow
   dim cell as tablecell
 response.write(rd.fieldcount)
    row = new tablerow()
	row.backcolor = drawing.color.yellow
     for i = 0 to rd.fieldcount - 1
	  cell = new tablecell()
	  cell.text = rd.getname(i)
      row.cells.add(cell)
     next
	table1.rows.add(row)
	while rd.read()
	 row = new tablerow()
	 for i = 0 to rd.fieldcount - 1
	  cell = new tablecell()
	  cell.text = rd.item(i)
	  row.cells.add(cell)
	 next
	 table1.rows.add(row)
	 end while
	end sub
</script>    
</head>
<body>
<form name='form1' id="form1" runat="server">
    <asp:table ID="table1" runat="server" />    
</form>
</body>
</html>
