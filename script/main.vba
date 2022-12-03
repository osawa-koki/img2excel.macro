Option Explicit

Sub Main()
  ' ファイルの読み込み
  Dim filename_raw As Variant
  filename_raw = Application.GetOpenFilename()
  If filename_raw = False Then
    Exit Sub
  End If
  Dim filename As String
  filename = filename_raw
  Set filename_raw = Nothing ' メモリの解放
  Debug.Print(filename)

  ' フリーファイル
  Dim free_file As Integer
  free_file = FreeFile()

  ' バイト格納用の配列
  Dim bytes() As Byte ' バイト配列

  ' ファイルをバイト配列に読み込む
  Open filename For Binary As #free_file ' バイナリモードでファイルを開く
    ReDim bytes(LOF(free_file)) ' バイト配列のサイズをセット
    Get #free_file, 1, bytes ' バイト配列にデータを格納
  Close #free_file
  Debug.Print(UBound(bytes))

End Sub
