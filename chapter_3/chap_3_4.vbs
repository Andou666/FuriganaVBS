' �t���ŌJ��Ԃ�
Dim i
Dim text, displayMsg
text = "�n���[�I" & vbCrLf

For i = 10 to 1 Step -1
    displayMsg = displayMsg & i & "��ڂ�" & text
Next

' 10������b�Z�[�W�{�b�N�X���o�������Ȃ��̂ňꊇ�ŕ\��������
' �\���F�un ��ڂ̃n���[�I�v
MsgBox displayMsg

