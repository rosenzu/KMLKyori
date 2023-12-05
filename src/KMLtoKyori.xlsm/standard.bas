Attribute VB_Name = "standard"
Option Explicit

'--------------------------------------------------
' 引数で指定したシートを削除
'--------------------------------------------------
Sub deleteWs(LA_strSeetName)
    On Error GoTo err_deleteWs
    
    Dim L_wsSeekSeets       As Worksheet
    
    'シート削除
    For Each L_wsSeekSeets In ThisWorkbook.Worksheets
        If L_wsSeekSeets.Name = LA_strSeetName Then
            L_wsSeekSeets.Delete
             
        End If
    Next
    
    Exit Sub

err_deleteWs:
    MsgBox "deleteWs ★エラー発生  " & Error

End Sub

'--------------------------------------------------
' VLookupで左側の列を探す（ワークシート内で使用）
'
' 例
'  =VLOOKUPLEFT("駅前",'06停留所'!B2:C9999,'06停留所'!C2:C9999,1)
'--------------------------------------------------
Public Function VLOOKUPLEFT(検索値 As Variant, データ範囲 As Variant, 検索範囲 As Variant, 返却列番号 As Integer) As Variant
On Error GoTo err_VLOOKUPLEFT

    VLOOKUPLEFT = WorksheetFunction.Index(データ範囲, WorksheetFunction.Match(検索値, 検索範囲, 0), 返却列番号)

    Exit Function
    
err_VLOOKUPLEFT:

    VLOOKUPLEFT = "XX"
    
End Function

Function std最終行(sname As String, Optional retsu As Long = 1) As Long
    
    std最終行 = Sheets(sname).Cells(Sheets(sname).Rows.Count, retsu).End(xlUp).Row

End Function

Function std最終列(sname As String, Optional gyou As Long = 1, Optional retsu As Long = 1) As Integer
    
    std最終列 = Sheets(sname).Cells(gyou, retsu).End(xlToRight).Column

End Function

'--------------------------------------------------
' ワークシート指定
' ワークブックが異なる場合はこれを使用する
'--------------------------------------------------

Function stdWs最終行(ws As Worksheet, Optional retsu As Long = 1) As Long
    
    stdWs最終行 = ws.Cells(ws.Rows.Count, retsu).End(xlUp).Row

End Function

'--------------------------------------------------
' ワークシート指定
' ワークブックが異なる場合はこれを使用する
'--------------------------------------------------

Function stdWs最終列(ws As Worksheet, Optional gyou As Long = 1, Optional retsu As Long = 1) As Integer
    
    stdWs最終列 = ws.Cells(gyou, retsu).End(xlToRight).Column

End Function

'--------------------------------------------------
' onedriveを私用している場合に、ローカルPathを取得
' 参考サイト　https://scodebank.com/?p=696
'--------------------------------------------------

Function UrlToLocal(ByRef Url As String) As String

   'OneDrive環境変数を格納する変数の定義
    Dim OneDrive As String

   'OneDrive環境変数の取得
    OneDrive = Environ("OneDrive")

   '「https://･･･････/Documents」までの文字数を格納する変数の定義
    Dim CharPosi As String

   ' ＵＲＬからローカルパスを作成する
    If Url Like "https://*" Then 'OneDriveのパスかどうかの判定
           
      '「https://･･･････/Documents」までの文字数を取得
      CharPosi = InStr(1, Url, "/Documents")
      
      'ローカルパス作成
      Url = OneDrive & Replace(Mid(Url, CharPosi), "/", Application.PathSeparator)
    
    Else
    
      'OneDriveのパス以外だったらカレントドライブ指定
      ChDrive Left(Url, 1)
     
    End If

  '作成したローカルパスを返す
   UrlToLocal = Url

End Function
