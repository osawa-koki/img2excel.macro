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

  ' ファイルフォーマットの判定
  Dim file_format As String
  file_format = Chr(HexToDec(bytes, 0, 0)) & Chr(HexToDec(bytes, 1, 1))
  If file_format <> "BM" Then
    MsgBox("Bitmapファイルを選択してください。")
    Exit Sub
  End If

  ' ファイルのヘッダ取得
  Dim width As Long '画像の横サイズ
  width = HexToDec(bytes, 18, 21)
  Dim height As Long '画像の縦サイズ
  height = HexToDec(bytes, 22, 25)

  Debug.Print("width -> " & CStr(width))
  Debug.Print("height -> " & CStr(height))

End Sub

' 連続したバイト配列の値を10進数に変換する関数
Function HexToDec(ByRef databuf, start, finish) As Long
  Dim i As Long
  Dim temp As String ' 16進数を格納する文字列配列
  temp = ""
  For i = finish To start Step -1 ' 後ろから処理
    temp = temp + Right("00" & Hex(databuf(i)), 2) ' 10進数を16進数に変換
  Next
  HexToDec = Val("&H" & temp) ' 16進数を10進数に変換
End Function
