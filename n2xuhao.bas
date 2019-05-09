Attribute VB_Name = "n2xuhao"
'行数转序号统计_助理V20190509w
'jianpf_20190509_w
'372191242@qq.com

'格式如下
'序号    收寄日期
'1       2019-04-03
'2       2019-04-03
'3       2019-04-03
'统计值3 空值识别位置断层
'空值 识别结束

Sub myxuhao()
Attribute myxuhao.VB_Description = "1行数转序号1234"
Attribute myxuhao.VB_ProcData.VB_Invoke_Func = "e\n14"

Dim kt开头序号值 As String
Dim kt开头日期值 As String
Dim dq当前序号值 As Integer
dq当前序号值 = 2
 

kt开头序号值 = dqh当前行序号值()
kt开头日期值 = dqh当前行日期值()

Debug.Print kt开头序号值; kt开头日期值

If kt开头序号值 <> "1" Then
msg = MsgBox("请选择开头序号位置: 1 的位置", vbOKOnly, "开头位置")
Exit Sub
End If
Debug.Print kt开头序号值

'选择下一行
'判断日期行是否为空
'循环处理
Do
 xzx选择下一行
kt开头序号值 = dqh当前行序号值()
kt开头日期值 = dqh当前行日期值()

'判断是否断层结束
If kt开头序号值 = "" Then
 msg = MsgBox("序号为空值,出现断层,请手动确认是否结束或重新开始!", vbOKOnly, "序号断层")
 xzx选择上一行
szd设置当前位置值 ("")
'szd设置当前位置值 (" 本程序有jianpf20190509w提供QQ:372191242")
 Exit Sub
End If

'判断是否到了统计位置
If kt开头日期值 = "" Then
Debug.Print "结束1个序号"
'重置统计位置值
szd设置当前位置值 (dq当前序号值 - 1)
'复位序号为1
dq当前序号值 = 1
'开始下1个序列
 xzx选择下一行
End If

szd设置当前位置值 (dq当前序号值)
Debug.Print "当前序号值:"; dq当前序号值; Chr(10) & Chr(13)
dq当前序号值 = dq当前序号值 + 1

Loop While kt开头序号值 <> ""


End Sub

Function dqh当前行序号值() As String

dqh当前行序号值 = ActiveCell.Value

End Function

Function dqh当前行日期值() As String

dqh当前行日期值 = ActiveCell.Offset(0, 1).Value

End Function

Function xzx选择下一行()
ActiveCell.Offset(1, 0).Select
End Function
Function xzx选择上一行()
ActiveCell.Offset(-1, 0).Select
End Function

Function szd设置当前位置值(n As String)
ActiveCell.Value = n
End Function

Function kbs拷贝上一行值()
ActiveCell.Value = ActiveCell.Offset(-1, 0).Value
End Function

