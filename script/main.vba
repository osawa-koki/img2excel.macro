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

  ' ベースネームを取得
  Dim sheet_name As String
  sheet_name = CreateObject("Scripting.FileSystemObject").GetBaseName(filename)
  Debug.Print("ベースネーム -> " & sheet_name)

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

  ' ヘッダサイズの取得
  Dim header_size As Integer
  header_size = HexToDec(bytes, 10, 13)
  Debug.Print("header_size -> " & CStr(header_size))

  ' シートの削除
  Application.DisplayAlerts = False ' メッセージを非表示
  Dim ws As Worksheet
  For Each ws In Worksheets
    If ws.Name = sheet_name Then
      ws.Delete
    End If
  Next ws
  Application.DisplayAlerts = True  ' メッセージを表示

  ' シートの追加
  Dim sheet As Worksheet
  Set sheet = Worksheets.Add
  sheet.Name = sheet_name
  sheet.Activate

  ' 行と列のサイズを設定
  Dim pixel_size As Integer
  pixel_size = 2
  sheet.Range(Rows(1), Rows(height + 1)).RowHeight = pixel_size * 0.75
  sheet.Range(Columns(1), Columns(width + 1)).ColumnWidth = pixel_size * 0.0594


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
