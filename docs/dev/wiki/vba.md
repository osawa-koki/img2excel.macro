# VBAに関する説明イロイロ

VBAの関数とか構文とか知らなかったこと、備忘録として。  

## ファイル番号

ファイルを管理するための番号で、一意である必要がある。  

ファイル番号の生成にはFreeFile関数が使用される。  
デフォルトで1以上255以下(引数が0)の整数が返却され、引数に1を指定すると256以上511以下の整数が返される。  

## LOF関数

Openステートメントを使用して開いたファイルのサイズをバイト単位で表す長整数型(Long)の値を返します。[^1]  

[^1]: <https://learn.microsoft.com/ja-jp/office/vba/language/reference/user-interface-help/lof-function>
