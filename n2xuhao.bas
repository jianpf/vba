Attribute VB_Name = "n2xuhao"
'����ת���ͳ��_����V20190509w
'jianpf_20190509_w
'372191242@qq.com

'��ʽ����
'���    �ռ�����
'1       2019-04-03
'2       2019-04-03
'3       2019-04-03
'ͳ��ֵ3 ��ֵʶ��λ�öϲ�
'��ֵ ʶ�����

Sub myxuhao()
Attribute myxuhao.VB_Description = "1����ת���1234"
Attribute myxuhao.VB_ProcData.VB_Invoke_Func = "e\n14"

Dim kt��ͷ���ֵ As String
Dim kt��ͷ����ֵ As String
Dim dq��ǰ���ֵ As Integer
dq��ǰ���ֵ = 2
 

kt��ͷ���ֵ = dqh��ǰ�����ֵ()
kt��ͷ����ֵ = dqh��ǰ������ֵ()

Debug.Print kt��ͷ���ֵ; kt��ͷ����ֵ

If kt��ͷ���ֵ <> "1" Then
msg = MsgBox("��ѡ��ͷ���λ��: 1 ��λ��", vbOKOnly, "��ͷλ��")
Exit Sub
End If
Debug.Print kt��ͷ���ֵ

'ѡ����һ��
'�ж��������Ƿ�Ϊ��
'ѭ������
Do
 xzxѡ����һ��
kt��ͷ���ֵ = dqh��ǰ�����ֵ()
kt��ͷ����ֵ = dqh��ǰ������ֵ()

'�ж��Ƿ�ϲ����
If kt��ͷ���ֵ = "" Then
 msg = MsgBox("���Ϊ��ֵ,���ֶϲ�,���ֶ�ȷ���Ƿ���������¿�ʼ!", vbOKOnly, "��Ŷϲ�")
 xzxѡ����һ��
szd���õ�ǰλ��ֵ ("")
'szd���õ�ǰλ��ֵ (" ��������jianpf20190509w�ṩQQ:372191242")
 Exit Sub
End If

'�ж��Ƿ���ͳ��λ��
If kt��ͷ����ֵ = "" Then
Debug.Print "����1�����"
'����ͳ��λ��ֵ
szd���õ�ǰλ��ֵ (dq��ǰ���ֵ - 1)
'��λ���Ϊ1
dq��ǰ���ֵ = 1
'��ʼ��1������
 xzxѡ����һ��
End If

szd���õ�ǰλ��ֵ (dq��ǰ���ֵ)
Debug.Print "��ǰ���ֵ:"; dq��ǰ���ֵ; Chr(10) & Chr(13)
dq��ǰ���ֵ = dq��ǰ���ֵ + 1

Loop While kt��ͷ���ֵ <> ""


End Sub

Function dqh��ǰ�����ֵ() As String

dqh��ǰ�����ֵ = ActiveCell.Value

End Function

Function dqh��ǰ������ֵ() As String

dqh��ǰ������ֵ = ActiveCell.Offset(0, 1).Value

End Function

Function xzxѡ����һ��()
ActiveCell.Offset(1, 0).Select
End Function
Function xzxѡ����һ��()
ActiveCell.Offset(-1, 0).Select
End Function

Function szd���õ�ǰλ��ֵ(n As String)
ActiveCell.Value = n
End Function

Function kbs������һ��ֵ()
ActiveCell.Value = ActiveCell.Offset(-1, 0).Value
End Function

