VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TimeCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'������̏o�Γ��A�o�Ύ��ԁA�ދΎ��Ԃ�\���N���X
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Public ���t As Date
Public �o�Ύ��� As String
Public �ދΎ��� As String

'���̓��̘J�����Ԃ������o�ϐ�����v�Z����B
'IsNumeric�Łu���x�v���̐����ȊO�̕����������Ă��Ȃ����`�F�b�N
'�[��J���̌v�Z:�[��ɏo�΂���ƃ}�C�i�X�ɂȂ�̂�24�����Ă���
'�x�e���Ԃ̌v�Z:�W���Ԉȏ㓭������P���ԋx��ł���͂��Ƃ����v�Z
Public Property Get �J������() As Integer
    If IsNumeric(�o�Ύ���) And IsNumeric(�ދΎ���) Then
        �J������ = �ދΎ��� - �o�Ύ���
        If �J������ < -1 Then
            �J������ = �J������ + 24
        End If
        If �J������ >= 8 Then
            �J������ = �J������ - 1
        End If
    Else
    �J������ = 0
    End If
End Property