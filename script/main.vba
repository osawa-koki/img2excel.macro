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
  Dim pixel_count As Long '画像のピクセル数
  pixel_count = width * height
  Debug.Print("width -> " & CStr(width))
  Debug.Print("height -> " & CStr(height))
  Debug.Print("pixel_count -> " & CStr(pixel_count))

  ' ヘッダサイズの取得
  Dim header_size As Integer
  header_size = HexToDec(bytes, 10, 13)
  Debug.Print("header_size -> " & CStr(header_size))

  ' カラーパレットの取得
  Dim color_palette As Integer
  color_palette = HexToDec(bytes, 28, 29)
  Debug.Print("color_palette -> " & CStr(color_palette))
  Select Case color_palette
    Case 24
      color_palette = 3
    Case 32
      color_palette = 4
    Case Else
      MsgBox("カラーパレットが24bitまたは32bitではありません。")
      Exit Sub
  End Select

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
  sheet.Range(Rows(1), Rows(height)).RowHeight = 7.5
  sheet.Range(Columns(1), Columns(width)).ColumnWidth = 0.77

  ' ピクセルデータの作成
  Dim colors() As PixelInfo ' 色配列
  ReDim colors(pixel_count) ' 色配列のサイズをセット
  Dim pixel_counter As Integer
  Dim pixel_total As Long
  pixel_total = UBound(bytes)
  For pixel_counter = header_size To pixel_total - 1 Step color_palette
    Dim B As Integer
    Dim G As Integer
    Dim R As Integer
    B = HexToDec(bytes, pixel_counter, pixel_counter)
    G = HexToDec(bytes, pixel_counter + 1, pixel_counter + 1)
    R = HexToDec(bytes, pixel_counter + 2, pixel_counter + 2)
    Dim index As Integer
    index = (pixel_counter - header_size) / color_palette
    Set colors(index) = New PixelInfo
    colors(index).Color = RGB(R, G, B)
    colors(index).X = index Mod width
    colors(index).Y = height - index \ width - 1
  Next pixel_counter

  ' ランダム配列
  Dim randoms() As Integer
  ReDim randoms(pixel_count)
  Dim random_counter As Integer
  For random_counter = 0 To pixel_count - 1
    randoms(random_counter) = random_counter
  Next random_counter
  Dim random_index As Integer
  For random_index = 0 To pixel_count - 1
    Dim random As Integer
    random = Int(Rnd() * pixel_count)
    Dim temp As Integer
    temp = randoms(random_index)
    randoms(random_index) = randoms(random)
    randoms(random) = temp
  Next random_index

  ' ピクセルデータの書き込み
  For pixel_counter = 0 To pixel_count - 1
    DIm x_index As Integer
    x_index = randoms(pixel_counter)
    Dim x As Integer
    x = colors(x_index).X
    Dim y As Integer
    y = colors(x_index).Y
    Dim color As Long
    color = colors(x_index).Color
    sheet.Cells(y + 1, x + 1).Interior.Color = color
  Next pixel_counter

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
