Option Explicit

'画像ファイルをシート上に展開するプログラム
Sub Main()

  Application.ScreenUpdating = False

  Dim filename As String 'ファイル名
    filename = Application.GetOpenFilename '読み込むファイル
  Dim ff As Long 'フリーファイル
    ff = FreeFile 'フリーファイル変数
  Dim databuf() As Byte 'バイト配列

  Open filename For Binary As #ff 'バイナリモードでファイルを開く
    ReDim databuf(LOF(ff)) 'バイト配列のサイズをセット
    Get #ff, 1, databuf 'バイト配列にデータを格納
  Close #ff

  '=====================================================================================
  Dim image_power As Long '画像の倍率
  '=====================================================================================
    image_power = 1  '倍率を1/image_powerで指定
  '=====================================================================================
  Dim bmp_width As Long '画像の横サイズ
    bmp_width = HexToDec(databuf, 18, 21) '4バイト
  Dim bmp_height As Long '画像の縦サイズ
    bmp_height = HexToDec(databuf, 22, 25) '4バイト
  Dim bmp_bit As Integer '画像のビットサイズ
    bmp_bit = HexToDec(databuf, 28, 29) '2バイト
  Dim file_size As Long 'ファイルサイズ
    file_size = HexToDec(databuf, 2, 5) '4バイト

  Cells.Clear 'セルをクリア

  Range(Columns(1), Columns(bmp_width / image_power)).ColumnWidth = 0.31 'セルの横幅
  Range(Rows(1), Rows(bmp_height / image_power)).RowHeight = 3 'セルの縦幅

  Dim width_size As Long 'セル上の横サイズ
    width_size = Fix(bmp_width / image_power)
  Dim height_size As Long 'セル上の縦サイズ
    height_size = Fix(bmp_height / image_power)
    If height_size = 0 Then
      MsgBox "指定した倍率が小さすぎます", vbExclamation
      Exit Sub '高さが0になる場合、終了
    End If

  '=====================================================================================
  '倍率を考慮して補正するバイト数を調整
  '=====================================================================================
  Dim widthcount As Long '横のデータ数
  Dim addpos As Integer '埋めるバイト数
  widthcount = Fix(bmp_width * 3)
  If widthcount Mod 4 > 0 Then '※widthが4の倍数に満たない場合、横の実データ数を求める
    widthcount = Fix(widthcount / 4 + 1) * 4
    '例
    'widthcount = 192 * 3 = 576 → widthcount Mod 4 = 0バイト埋める → 1列 = 576バイト
    'widthcount = 191 * 3 = 573 → widthcount Mod 4 = 3バイト埋める → 1列 = 576バイト
    'widthcount = 190 * 3 = 570 → widthcount Mod 4 = 2バイト埋める → 1列 = 572バイト
    'widthcount = 189 * 3 = 567 → widthcount Mod 4 = 1バイト埋める → 1列 = 568バイト
  End If
  addpos = widthcount - width_size * image_power * 3 '倍率に応じた不足分を調整
  '=====================================================================================

  Dim pos As Long 'データオフセット
  Dim w_index As Long 'ビットマップの横位置
  Dim h_index As Long 'ビットマップの縦位置
    w_index = 1 '横の初期位置座標(左)をセット
    h_index = height_size '縦の初期位置座標(下)をセット

  pos = 54 '初期値(データ先頭位置)
  For h_index = h_index To 1 Step -1 '高さのループ
    For w_index = 1 To width_size '幅のループ
      Cells(h_index, w_index).Interior.Color = _
        RGB(HexToDec(databuf, pos + 2, pos + 2), _
        HexToDec(databuf, pos + 1, pos + 1), _
        HexToDec(databuf, pos, pos)) 'データに対応するセル背景色にRGBを指定
      pos = pos + 3 * image_power '横移動
    Next
    pos = pos + addpos '4バイト区切りの不足分を加算(posを調整)
    pos = pos + widthcount * (image_power - 1) '列×倍率分飛ばす(縦移動)
  Next
  MsgBox "画像の展開が完了しました" '処理完了
End Sub

'連続したバイト配列の値を10進数に変換する関数
Function HexToDec(ByRef databuf, first, last) As Long
  Dim i As Long 'ループカウンタ
  Dim temp As String '16進数を格納する文字列配列
    temp = ""
  For i = last To first Step -1 '後ろから処理
    temp = temp + Right("00" & Hex(databuf(i)), 2) '10進数を16進数に変換
  Next
  HexToDec = Val("&H" & temp) '16進数を10進数に変換
End Function

