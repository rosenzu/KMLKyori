Attribute VB_Name = "Module1"
' KLM->�����v�Z�c�[��
' ������ʗ��p���i�l�b�g���[�N�@�ɓ��_�V

Option Explicit

Sub ClearSheet()
    
    Dim L_wsMenu    As Worksheet
    Set L_wsMenu = Sheets("���j���[")

    Dim L_wsKml     As Worksheet
    Set L_wsKml = Sheets("KML")
    
    Dim L_wsShapes  As Worksheet
    Set L_wsShapes = Sheets("Shapes")
    
    
    '�V�[�g�N���A
    L_wsKml.Cells.Clear
    L_wsKml.Activate
    L_wsKml.Range("A1").Activate

    '�V�[�g�N���A
    L_wsShapes.Cells.Clear
    L_wsShapes.Activate
    L_wsShapes.Range("A1").Activate


    L_wsMenu.Activate
    MsgBox "�V�[�g�̃N���A���������܂����B"

End Sub

Sub ImportKml()
    On Error GoTo err_ImportKml
    
    Dim WshShell As Object
    Set WshShell = CreateObject("WScript.Shell")
    
    Dim L_strOpenFilePath As String
    
    Dim L_lngRow    As Long
    Dim L_strReadString As String
    
    Dim L_objStream As Object
    Set L_objStream = CreateObject("ADODB.Stream")
    
    
    Dim L_wsKml     As Worksheet
    Set L_wsKml = Sheets("KML")
    
    '------
    '�J�����g�t�H���_�̐ݒ�
    WshShell.CurrentDirectory = UrlToLocal(ActiveWorkbook.Path)
    
    '------
    '�_�C�A���O�\�����āA���[�U�[���I�������t�H���_���擾
    L_strOpenFilePath = Application.GetOpenFilename("�e�L�X�g�t�@�C��,*.kml?", 1, "KML�t�@�C����I�����Ă��������B")
    
    
    '------
    '�e�L�X�g�`����KML����݂���
    'LF�ł��ǂݍ��߂�悤 ADODB���g�p
    '�Q�l�T�C�g�@Excel�FVBA�FUTF-8�^LF�̃t�@�C����ǂݍ���
    'http://www.hiihah.info/index.php?Excel%EF%BC%9AVBA%EF%BC%9AUTF-8%EF%BC%8FLF%E3%81%AE%E3%83%95%E3%82%A1%E3%82%A4%E3%83%AB%E3%82%92%E8%AA%AD%E3%81%BF%E8%BE%BC%E3%82%80
     
    L_objStream.Type = 2         '������^
    L_objStream.Charset = "utf-8"
    L_objStream.LineSeparator = 10
    L_objStream.Open
    
    L_objStream.LoadFromFile (L_strOpenFilePath)
  
    L_lngRow = stdWs�ŏI�s(L_wsKml) + 1
    
    
    Do While Not L_objStream.EOS
        L_strReadString = L_objStream.ReadText(-2)           '�e�L�X�g��1�s�ǂݍ��ށB
        L_wsKml.Cells(L_lngRow, 1).Value = L_strReadString
        L_lngRow = L_lngRow + 1
    Loop
 
    L_objStream.Close
    Set L_objStream = Nothing
     
    MsgBox "KML�t�@�C���̓ǂݍ��݂��������܂���"
    
     
    Exit Sub
    
err_ImportKml:
    If Err = 3002 Then
        MsgBox "�����𒆎~���܂��B"
    Else
        MsgBox "ImportKml ���G���[����  " & Error
    End If

End Sub

Sub run_makeKyori()
    
    Call makeKyori
    
    MsgBox "�����̏o�͂��������܂���"

End Sub


Sub makeKyori()
    On Error GoTo err_makeShapes
    
    Dim WshShell As Object
    Set WshShell = CreateObject("WScript.Shell")
    
    Dim L_strOpenFilePath As String
    
    Dim L_lngRow    As Long
    Dim L_strReadString As String
    
    Dim L_objStream As Object
    Set L_objStream = CreateObject("ADODB.Stream")
    
    
    Dim L_wsMenu     As Worksheet
    Set L_wsMenu = Sheets("���j���[")
    
    Dim L_wsKml     As Worksheet
    Set L_wsKml = Sheets("KML")
    
    Dim L_lngRowKml As Long
    Dim L_lngMaxRowKML As Long
    
    Dim L_wsShapes  As Worksheet
    Set L_wsShapes = Sheets("Shapes")
    
    Dim L_lngRowShapes As Long
    
    Dim L_lngLoopCount As Long
    
    Dim L_strKmlValue As String
    Dim L_strName As String
    Dim L_strFlolderName As String
    Dim L_strNameBfr As String          'Name�ޔ�p
    Dim L_blnFolder As Boolean          'Folder�����t���O
    Dim L_blnPlacemark As Boolean       'Placemaak�����t���O
    Dim L_blnLineString As Boolean      'LineString�����t���O
    Dim L_blnCoordinates As Boolean     'coordinates�����t���O
    Dim L_blnDouble As Boolean
    Dim L_lngSequence As Long
    
    Dim L_vntSplit As Variant
    Dim L_vntSplit2 As Variant
    Dim L_strLat As String
    Dim L_strLon As String
    Dim L_strLatBfr As String   '�ޔ�p
    Dim L_strLonBfr As String   '�ޔ�p
    
    '------
    '�w�b�_�쐬
    L_lngRowShapes = 1
    L_wsShapes.Cells(L_lngRowShapes, 1) = "shape_id"
    L_wsShapes.Cells(L_lngRowShapes, 2) = "shape_pt_lat"
    L_wsShapes.Cells(L_lngRowShapes, 3) = "shape_pt_lon"
    L_wsShapes.Cells(L_lngRowShapes, 4) = "shape_pt_sequence"
    L_wsShapes.Cells(L_lngRowShapes, 5) = "kyori"
    
    '------
    'KML�V�[�g�@�s���[�v
    L_lngMaxRowKML = stdWs�ŏI�s(L_wsKml)
    
    
    L_blnFolder = False
    L_blnPlacemark = False
    L_blnLineString = False
    L_blnCoordinates = False
    
    L_strNameBfr = ""
    L_strLatBfr = ""
    L_strLonBfr = ""
    L_lngSequence = 1
    
    For L_lngRowKml = 1 To L_lngMaxRowKML
        
        L_strKmlValue = L_wsKml.Cells(L_lngRowKml, 1).Value
        L_strKmlValue = Replace(L_strKmlValue, vbTab, "")
        L_strKmlValue = Trim(L_strKmlValue)
        
        If InStr(L_strKmlValue, "Folder") And L_blnFolder = False Then
            L_blnFolder = True
        End If
        
        If InStr(L_strKmlValue, "/Folder") And L_blnFolder = True Then
            L_blnFolder = False
            L_strName = ""
        End If
        
        If InStr(L_strKmlValue, "Placemark") And L_blnPlacemark = False Then
            L_blnPlacemark = True
        End If
    
        If InStr(L_strKmlValue, "/Placemark") And L_blnPlacemark = True Then
            L_blnPlacemark = False
            L_strName = ""
        End If
    
        If InStr(L_strKmlValue, "LineString") And L_blnPlacemark = True And L_blnLineString = False Then
            L_blnLineString = True
        End If
        
        If InStr(L_strKmlValue, "/LineString") And L_blnPlacemark = True And L_blnLineString = True Then
            L_blnLineString = False
            L_strName = ""
        End If
        
        If InStr(L_strKmlValue, "coordinates") And L_blnPlacemark = True And L_blnLineString = True And L_blnCoordinates = False Then
            L_blnCoordinates = True
        End If
        
        If InStr(L_strKmlValue, "/coordinates") And L_blnPlacemark = True And L_blnLineString = True And L_blnCoordinates = True Then
            L_blnCoordinates = False
            L_strName = ""
        End If
        
        
        'Folder�v�f�̏ꍇ
        If (L_wsMenu.Range("route_id_B") = "#") Or (L_wsMenu.Range("route_id_B") = "��") Then
            If L_blnFolder = True And L_blnPlacemark = False Then
                
                'name�v�f�̏ꍇname���擾
                If InStr(L_strKmlValue, "name") Then
                    L_strFlolderName = Replace(L_strKmlValue, "<name>", "")
                    L_strFlolderName = Replace(L_strFlolderName, "</name>", "")
                    
                    'name��"_"���܂ޏꍇ�A"_"�ȑO���̗p
                    If InStr(L_strFlolderName, "_") >= 1 Then
                        L_vntSplit = Split(L_strFlolderName, "_")
                        L_strFlolderName = L_vntSplit(0)
                    End If
                    
                    L_strName = L_strFlolderName
                End If
            End If
        End If
        
        If L_blnFolder = False Then
            L_strFlolderName = ""
        End If
        
        
        'Placemark�v�f�̏ꍇ
        If L_blnPlacemark = True Then
            
            If (L_wsMenu.Range("route_id_A") = "#") Or (L_wsMenu.Range("route_id_A") = "��") Then
                'name�v�f�̏ꍇname���擾 (Folder��name���擾����Ă��Ȃ��ꍇ)
                If InStr(L_strKmlValue, "name") And L_strFlolderName = "" And L_strName = "" Then
                    L_strName = Replace(L_strKmlValue, "<name>", "")
                    L_strName = Replace(L_strName, "</name>", "")
                    
                    'name��"_"���܂ޏꍇ�A"_"�ȑO���̗p
                    If InStr(L_strName, "_") Then
                        L_vntSplit = Split(L_strName, "_")
                        L_strName = L_vntSplit(0)
                    End If
                End If
            End If
                
                
            '���C���̍��W�̏ꍇ�Ashapes�V�[�g�̏o��
            If L_blnCoordinates Then
            
                If InStr(L_strKmlValue, "<") = False Then
                
                    'KML�̈ܓx�o�x
                    
                    '�X�y�[�X��؂�̏ꍇ�@�J��Ԃ�
                    If InStr(L_strKmlValue, " ") >= 1 Then
                        L_vntSplit = Split(Trim(L_strKmlValue), " ")
                        
                    Else
                        ReDim L_vntSplit(0)
                        L_vntSplit(0) = Trim(L_strKmlValue)
                    End If
                        
                        
                    For L_lngLoopCount = 0 To UBound(L_vntSplit)
                        L_vntSplit2 = Split(Trim(L_vntSplit(L_lngLoopCount)), ",")
                        L_strLat = L_vntSplit2(1)
                        L_strLon = L_vntSplit2(0)
            
            
                        '�ޔ�����name�ƈقȂ�ꍇ�i�ٌn���j�́Asequence�����Z�b�g����
                        L_blnDouble = False
                        
                        If L_strName <> L_strNameBfr Then
                            L_lngSequence = 1
                        Else
                            
                            '���n���ŁA��O�̂ƈܓx�o�x�������ꍇ�͏o�͑ΏۊO�Ƃ���
                            If L_strLat = L_strLatBfr And L_strLon = L_strLonBfr Then
                                L_blnDouble = True
                            Else
                                L_lngSequence = L_lngSequence + 1
                            End If
                        End If
                        
                        If L_blnDouble = False Then
                            L_lngRowShapes = L_lngRowShapes + 1
                        
                            L_wsShapes.Cells(L_lngRowShapes, 1) = L_strName
                            L_wsShapes.Cells(L_lngRowShapes, 2) = L_strLat
                            L_wsShapes.Cells(L_lngRowShapes, 3) = L_strLon
                            L_wsShapes.Cells(L_lngRowShapes, 4) = L_lngSequence
                            
                            '�����v�Z�� �V�[�P���X2�ȏ�̏ꍇ
                            If L_lngSequence >= 2 Then
                                '�����v�Z�@�Q�ƃT�C�g https://zenn.dev/music_shio/articles/3c59e10842fcc7
                                L_wsShapes.Cells(L_lngRowShapes, 5).FormulaR1C1 = "=SQRT(((R[-1]C[-3] - RC[-3])*PI()/180*6378137*(1-0.00669437999)/SQRT(1-0.00669437999*SIN((R[-1]C[-3] + RC[-3])/2*PI()/180)^2)^3)^2+((R[-1]C[-2] - RC[-2])*PI()/180*6378137/SQRT(1-0.00669437999*SIN((R[-1]C[-3] + RC[-3])/2*PI()/180)^2)*COS((R[-1]C[-3] + RC[-3])/2*PI()/180))^2)"
                            End If
                        End If
                        
                        '�ޔ�����
                        L_strNameBfr = L_strName
                        L_strLatBfr = L_strLat
                        L_strLonBfr = L_strLon
                
                    Next L_lngLoopCount
                
                End If
            End If
    
        End If
    Next L_lngRowKml
    
    'L_wsShapes.Activate
    'L_wsShapes.Range("A1").Activate
    
    '�s�{�b�g�e�[�u�����X�V
    Worksheets("�n���ʋ���").PivotTables("�n���ʋ���PT").PivotCache.Refresh
    
    
    Exit Sub
    
err_makeShapes:
    MsgBox "makeShapes ���G���[����  " & Error

End Sub
