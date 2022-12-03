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
End Sub
