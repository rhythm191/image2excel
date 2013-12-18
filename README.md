images2excel
============

ディレクトリに格納された画像群を貼り付けた状態のExcelファイルを生成するrubyスクリプト.


Configration
--------

* 対象となる画像群はimgディレクトリに格納する(config.ymlで変更可能)
* エクスポートされるExcelファイルは"export.xls"(config.ymlで変更可能)

    .
    |-- images2excel.rb
    |-- export.xls
    `-- img
        |-- 1-1.png
        |-- 1-2.png
        `-- 1-3.png

対象となる画像形式は次の通り

* bmp (拡張子が bmp, BMP)
* jpeg (拡張子が jpg, jpeg, JPG)
* png  (拡張子が png, PNG)


Config File
-----------

config.ymlに動作の設定を記述する.
次の設定値が存在する.

* directory: 対象となる画像群の格納ディレクトリ(default: "img")
* export: エクスポートされるExcel(default: "export.xls")
* scale: 画像の拡大率(default: 0.65)
* space: 同一シートに複数の画像を配置する場合の画像間のスペース(default: 50)



Usage
-----

次のgemをインストール

```sh
gem install imagesize
```

またはbundleコマンドを使う

```sh
bundle install
```

次のコマンドを実行(rubyにパスが通っていればダブルクリックでも動作可能)

```sh
ruby images2excel.rb
```


Detail
------

画像のファイル名をそれぞれ `1-1.png, 1-2.png, 1-3.png` とした場合、
Excelに3つのシート( `1-1, 1-2, 1-3` )が作成され、それぞれに画像が貼り付けられた状態でエクスポートされる.

画像のファイル名をコンマがあった場合、例えば、 `1-1.01.png, 1-1.03.png, 1-1.03.png` とした場合、
Excelに1つのシート( `1-1` )が作成され、それぞれに画像が貼り付けられた状態でエクスポートされる.

上記のルールは混在してもいい.例えば、`1-1.png, 1-2.png, 1-3.00.png, 1-3.01.png` とした場合、
Excelに3つのシート( `1-1, 1-2, 1-3` )が作成され、それぞれに画像が貼り付けられた状態、
ただし、1-3にはふたつの画像が貼り付けられた状態でエクスポートされる.


シートはファイル名の先頭の数値の昇順に作られる。


また、'_'(アンダーバー)から始まる文字列をコメントとして認識します。

例えば、 ファイル名が'1-1._hogehoge.png'の場合、
'hogehoge'をコメントと認識し、画像の上部のセルに'hogehoge'が入力されます。

改行が必要な場合は'_'を使います。
例えば、 ファイル名が'1-1._hoge_fuga.png'の場合、
'hoge\nfuga'をコメントと認識し、画像の上部のセルに'fuga'が、その上のセルに'hoge'が入力されます。



Environment
-----------

次の環境での動作確認を行っています.

* windows 7
* excel 2003
* ruby 1.9.3-p374

