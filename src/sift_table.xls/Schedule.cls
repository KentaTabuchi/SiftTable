VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Schedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'�\���\���N���X
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Enum ��
    �J�n�� = 3
    �ŏI�� = 39
End Enum
Private Enum �s
    ���t�s = 4
End Enum
Public ��ƍs As Integer

'�f�[�^�x�[�X����\��̃��R�[�h�Z�b�g�����o���R���N�V�����Ɋi�[����
Public Property Get ��c��() As Collection
    Set ��c�� = New Collection
    DatabaseIO.DBConnect
    DatabaseIO.OpenAdo ("SELECT * FROM �\��")
    Dim recordCount As Integer
    recordCount = DatabaseIO.adoRs.recordCount
    If recordCount = 0 Then
        Debug.Print "�\�肪����܂���"
    Else
        Dim i As Integer
        For i = 0 To recordCount - 1
        On Error Resume Next 'NULL����������Ǝ~�܂邪�����Ă��󕶎�����������̂Ŗ���
            Dim �C�x���g As Events
            Set �C�x���g = New Events
            �C�x���g.���t = DatabaseIO.adoRs!���t
            �C�x���g.���e = DatabaseIO.adoRs!��c��
            �C�x���g.���� = DatabaseIO.adoRs!���� - 1
            ��c��.Add �C�x���g
            adoRs.MoveNext
        Next i
     End If
    DatabaseIO.CloseAdo
    DatabaseIO.DBClose
End Property
'�f�[�^�x�[�X����\��̃��R�[�h�Z�b�g�����o���R���N�V�����Ɋi�[����
Public Property Get �j�����X�g() As Collection
    Set �j�����X�g = New Collection
    DatabaseIO.DBConnect
    DatabaseIO.OpenAdo ("SELECT * FROM �j��")
    Dim recordCount As Integer
    recordCount = DatabaseIO.adoRs.recordCount
    If recordCount = 0 Then
        Debug.Print "�j��������܂���"
    Else
        Dim i As Integer
        For i = 0 To recordCount - 1
            Dim �C�x���g As Events
            Set �C�x���g = New Events
            �C�x���g.���t = DatabaseIO.adoRs!���t
            �C�x���g.���e = DatabaseIO.adoRs!�j��
            �C�x���g.���� = 0
            �j�����X�g.Add �C�x���g
            adoRs.MoveNext
        Next i
     End If
    DatabaseIO.CloseAdo
    DatabaseIO.DBClose
End Property
