<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html>
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html;charset=utf-8">
    <title>BI RUN 3D 自動化報表</title>
	<style>
		th{
			background: #98f5ff;
		}
		th,td{
			border: 2px solid #3366ff;  /* 设置边框宽度和颜色 */
		}
		#number_th{
			width: 100px;
			height: 10px;
		}
		#nbsno_th,#project_th{
			width: 230px;
			height: 10px;
		}
		#mac_th, #result_th{
			width: 180px;
			height: 10px;
		}
		#status_th{
			width: 70px;
			height: 10px;
		}
		#date_th{
			width: 230px;
			height: 10px;
		}
		
		table {
		  border-collapse: collapse; /* 合并边框 */
		  text-align: center;
		  border: 2px solid #3366ff; /* 设置边框宽度和颜色 */
		}
    </style>
</head>
<body align="center">
     <table align="center" border="1">
        <tr>
			<th id='number_th'>序號</th>
            <th id='nbsno_th'>NBSNO</th>
            <th id='mac_th'>MAC地址</th>
            <th id='project_th'>測試項目</th>
            <th id='result_th'>測試結果</th>
            <th id='status_th'>測試狀態</th>
            <th id='date_th'>測試日期</th>
        </tr>
		<%
			Dim paramName, paramValue, number
			paramName = "device_name"
			paramValue = Request.QueryString(paramName)
			Response.ContentType = "text/html"
			Response.Charset = "utf-8"
			number = 1
			
			'建数据库连接对象
            Set conn = Server.CreateObject("ADODB.Connection")
            ' 设置连接字符串
            connStr = "Provider=SQLOLEDB;Data Source=BIOS2;Initial Catalog=SWPRODUCE;User ID=sa;Password=P@ssw0rd"
            conn.Open connStr
            
            ' 创建记录集对象
            Set rs = Server.CreateObject("ADODB.Recordset")
			' 执行SQL查询
			'sql = "select top 10 A.NBSNO,B.* from  NBSNOCHKLOG A,pd_test_data B where B.mac=A.MAC and  A.NBSNO='NKV250RNC1WK30004G00701'"
            sql = "select A.NBSNO,B.* from NBSNOCHKLOG A,pd_test_data B  where B.mac=A.MAC AND B.create_date>'2024-01-01' AND A.NBSNO like"
			sql = sql & "'%" & paramValue & "%' AND B.item_no IN ('firestrikecombinedscorep','firestrikegraphicsscorep','firestrikeoverallscorep','firestrikephysicsscorep') AND A.NBSNO not in (select A.NBSNO from NBSNOCHKLOG A,pd_test_data B  where B.mac=A.MAC  and A.NBSNO like"
			sql = sql & "'%" & paramValue & "%' AND B.item_no IN ('firestrikecombinedscorep','firestrikegraphicsscorep','firestrikeoverallscorep','firestrikephysicsscorep') group by A.NBSNO having count(B.item_no) < 4)"
            rs.Open sql, conn
			
            ' 处理查询结果
            Do While Not rs.EOF
                Response.Write "<tr>"
				Response.Write "<td>" & number & "</td>"
                Response.Write "<td>" & rs("NBSNO") & "</td>"
                Response.Write "<td>" & rs("mac") & "</td>"
				Response.Write "<td>" & rs("item_no") & "</td>"
				Response.Write "<td>" & rs("result_log") & "</td>"
                Response.Write "<td>" & rs("create_user") & "</td>"
                Response.Write "<td>" & rs("create_date") & "</td>"
				Response.Write "</tr>"
				number = number + 1
                rs.MoveNext
            Loop
            
            ' 关闭记录集和连接
            rs.Close
            conn.Close
            Set rs = Nothing
            Set conn = Nothing
        %>
    </table>
</body>
</html>
