VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TimeCardButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'���I�Ƀ{�^������邽�߂̃N���X�B
'TimeCardButtonForm�ɒ���t����{�^�����`����
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

Public WithEvents button As CommandButton
Attribute button.VB_VarHelpID = -1
Public �o�ދ� As TimeCard
'TimeCardButtonForm�̃{�^���������ꂽ�Ƃ��̏���
Private Sub button_Click()
    WorkSheetWriter.WriteTimeCard Me.�o�ދ�
End Sub
