Attribute VB_Name = "standard"
Option Explicit

'--------------------------------------------------
' �����Ŏw�肵���V�[�g���폜
'--------------------------------------------------
Sub deleteWs(LA_strSeetName)
    On Error GoTo err_deleteWs
    
    Dim L_wsSeekSeets       As Worksheet
    
    '�V�[�g�폜
    For Each L_wsSeekSeets In ThisWorkbook.Worksheets
        If L_wsSeekSeets.Name = LA_strSeetName Then
            L_wsSeekSeets.Delete
             
        End If
    Next
    
    Exit Sub

err_deleteWs:
    MsgBox "deleteWs ���G���[����  " & Error

End Sub

'--------------------------------------------------
' VLookup�ō����̗��T���i���[�N�V�[�g���Ŏg�p�j
'
' ��
'  =VLOOKUPLEFT("�w�O",'06�◯��'!B2:C9999,'06�◯��'!C2:C9999,1)
'--------------------------------------------------
Public Function VLOOKUPLEFT(�����l As Variant, �f�[�^�͈� As Variant, �����͈� As Variant, �ԋp��ԍ� As Integer) As Variant
On Error GoTo err_VLOOKUPLEFT

    VLOOKUPLEFT = WorksheetFunction.Index(�f�[�^�͈�, WorksheetFunction.Match(�����l, �����͈�, 0), �ԋp��ԍ�)

    Exit Function
    
err_VLOOKUPLEFT:

    VLOOKUPLEFT = "XX"
    
End Function

Function std�ŏI�s(sname As String, Optional retsu As Long = 1) As Long
    
    std�ŏI�s = Sheets(sname).Cells(Sheets(sname).Rows.Count, retsu).End(xlUp).Row

End Function

Function std�ŏI��(sname As String, Optional gyou As Long = 1, Optional retsu As Long = 1) As Integer
    
    std�ŏI�� = Sheets(sname).Cells(gyou, retsu).End(xlToRight).Column

End Function

'--------------------------------------------------
' ���[�N�V�[�g�w��
' ���[�N�u�b�N���قȂ�ꍇ�͂�����g�p����
'--------------------------------------------------

Function stdWs�ŏI�s(ws As Worksheet, Optional retsu As Long = 1) As Long
    
    stdWs�ŏI�s = ws.Cells(ws.Rows.Count, retsu).End(xlUp).Row

End Function

'--------------------------------------------------
' ���[�N�V�[�g�w��
' ���[�N�u�b�N���قȂ�ꍇ�͂�����g�p����
'--------------------------------------------------

Function stdWs�ŏI��(ws As Worksheet, Optional gyou As Long = 1, Optional retsu As Long = 1) As Integer
    
    stdWs�ŏI�� = ws.Cells(gyou, retsu).End(xlToRight).Column

End Function

'--------------------------------------------------
' onedrive�����p���Ă���ꍇ�ɁA���[�J��Path���擾
' �Q�l�T�C�g�@https://scodebank.com/?p=696
'--------------------------------------------------

Function UrlToLocal(ByRef Url As String) As String

   'OneDrive���ϐ����i�[����ϐ��̒�`
    Dim OneDrive As String

   'OneDrive���ϐ��̎擾
    OneDrive = Environ("OneDrive")

   '�uhttps://�������/Documents�v�܂ł̕��������i�[����ϐ��̒�`
    Dim CharPosi As String

   ' �t�q�k���烍�[�J���p�X���쐬����
    If Url Like "https://*" Then 'OneDrive�̃p�X���ǂ����̔���
           
      '�uhttps://�������/Documents�v�܂ł̕��������擾
      CharPosi = InStr(1, Url, "/Documents")
      
      '���[�J���p�X�쐬
      Url = OneDrive & Replace(Mid(Url, CharPosi), "/", Application.PathSeparator)
    
    Else
    
      'OneDrive�̃p�X�ȊO��������J�����g�h���C�u�w��
      ChDrive Left(Url, 1)
     
    End If

  '�쐬�������[�J���p�X��Ԃ�
   UrlToLocal = Url

End Function
