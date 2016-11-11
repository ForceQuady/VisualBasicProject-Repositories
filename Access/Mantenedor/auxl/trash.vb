
''
'Function IsValidUserAndPwd(strName, strPwd)
'    '
'    ' Use a trusted connection to SQL server. Never use the sa account. 

'  Set cn = CreateObject("ADODB.Connection") 
'  cn.Open strConn 

'  Set cmd = CreateObject("ADODB.Command") 
'  cmd.ActiveConnection = cn
'    cmd.CommandText =
'      "select * from MyTable where Name=? and Pwd=?"
'    cmd.CommandType = 1 'adCmdText 
'    cmd.Prepared = True

'' Explanation of numeric parameters: 
'' data type is 200, varchar string 
'' direction is 1, input parameter only 
'' size of data is 32 chars max. 
'Set parm1 = cmd.CreateParameter("Name", 200, 1, 32, "") 
'  cmd.Parameters.Append parm1 
'  parm1.Value = strName

'Set parm2 = cmd.CreateParameter("Pwd", 200, 1, 32, "") 
'  cmd.Parameters.Append parm2 
'  parm1.Value = strPwd

'Set rs = cmd.Execute 
'  IsValidUserAndPwd = False
'    If 1 = rs(0).value Then IsValidUserAndPwd = True

'    rs.Close
'    cn.Close

'End Function