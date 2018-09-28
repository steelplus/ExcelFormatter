# 本ツールについて

ディレクトリ内にあるExcelファイル(.xlsx|.xls)に対して、設定に従いフォーマットをかけるツールです。

# 使い方

```ExcelFormatter.exe <処理対象ディレクトリ名> <option>```

処理対象のディレクトリは複数指定可能です。

|option|内容|
|---|---|
|``-a1 --selecta1``|指定のシートに対して、「A1セル」を選択した状態にします。|
|``-te --template``|指定のシートに対してテンプレートを適用します。|
|``-pa --page``|指定のシートに対してページ設定を適用します。|
|``-fo --font``|指定のシートに対してフォント設定を適用します。|
|``-hd --header``|指定のシートに対してヘッダ設定を適用します。|
|``-ft --footer``|指定のシートに対してフッタ設定を適用します。|
|``-sr --removesample``|指定のシートに対してSAMPLE文字を取り除きます。|
|``-sa --addsample``|指定のシートに対してSAMPLE文字を追加します。|
|``-ts --timestamp``|対象ファイルのタイムスタンプを上書きします。|
|``-mr --mirror``|処理対象ファイルを上書きせず、`<output dir>`に処理後のファイルをコピーします。|
|``-sd --subDirectory``|処理対象について、サブフォルダも検索します。|
|``-c --config``|コンフィグファイル(``config.json``)の雛形を作成し、アプリケーションを終了します。|

実行ファイルと同ディレクトリに ``template.xlsx``と``config.json``を配置して下さい。

# 設定ファイルについて

``config.json``にて設定を保持しています。

各設定は以下の共通のプロパティを持ちます。

|プロパティ名|内容|
|---|---|
|``applyTemplateSheetRegex``|処理を実行するシートを正規表現で指定します。空文字の場合はすべてのシートに対して実行します。|
|``excludeTemplateSheetRegex``|処理を除外するシートを正規表現で指定します。空文字の場合は無効となります。|

## ``selectA1``
``selectA1``プロパティは``-a1``を実行する際の設定項目です。

## ``template``

``template``プロパティは``-te``を実行する際の設定項目です。

|プロパティ名|内容|
|---|---|
|``template.templateSheetNum``|``template.xlsx``ファイル内のテンプレートがあるシートを番号で指定します。|
|``template.applyTemplateSheetRegex``|テンプレートを適用するシートを正規表現で入力します。空文字の場合はすべてのシートに適用します。|
|``template.excludeTemplateSheetRegex``|テンプレートを適用しないシートを正規表現で入力します。空文字の場合は無効となります。|
|``template.templateCellFrom``|テンプレートがあるシートのどの位置からコピーするかを指定します。左上のセルを指定します。|
|``template.templateCellEnd``|テンプレートがあるシートのどの位置からコピーするかを指定します。右下のセルを指定します。|
|``template.templateVals``|テンプレートシートの他に、指定したセルに値を入力する場合は指定します。|
|``template.templateVals.cell``|指定セル|
|``template.templateVals.value``|入力する値|

## ``pageSetup``

``pageSetup``プロパティは``-pa``を実行する際の設定項目です。

|プロパティ名|内容|
|---|---|
|``pageSetup.topMargin``|印刷設定余白：上|
|``pageSetup.bottmonMargin``|印刷設定余白：下|
|``pageSetup.leftMargin``|印刷設定余白：左|
|``pageSetup.rightMargin``|印刷設定余白：右|
|``pageSetup.headerMargin``|印刷設定余白：ヘッダー|
|``pageSetup.footerMargin``|印刷設定余白：フッター|
|``pageSetup.zoom``|シート拡大率(%)|
|``pageSetup.paperSize``|印刷サイズ。[こちら](https://msdn.microsoft.com/ja-jp/vba/excel-vba/articles/xlpapersize-enumeration-excel)の「名前」を指定してください。|
|``pageSetup.orientation``|印刷の向き。[こちら](https://msdn.microsoft.com/ja-jp/vba/excel-vba/articles/xlpageorientation-enumeration-excel)の「名前を指定してください。|

## ``font``
``font``プロパティは``-fo``を実行する際の設定項目です。

|プロパティ名|内容|
|---|---|
|``font.fontName``|変更先のフォント名|

## ``header``
``header``プロパティは``-hd``を実行する際の設定項目です。

|プロパティ名|内容|
|---|---|
|``header.centerHeader``|ヘッダー中央|
|``header.leftHeader``|ヘッダー左|
|``header.rightHeader``|ヘッダー右|

## ``footer``
``footer``プロパティは``-hd``を実行する際の設定項目です。

|プロパティ名|内容|
|---|---|
|``footer.centerFooter``|フッター中央|
|``footer.leftFooter``|フッター左|
|``footer.rightFooter``|フッター右|

## ``removeSample``
``removeSample``プロパティは``-sr``を実行する際の設定項目です。

|プロパティ名|内容|
|---|---|
|``removeSample.text``|取り除くサンプル文字の文字情報。完全一致したものを取り除きます。|

## ``addSample``

``addSample``プロパティは``-sa``を実行する際の設定項目です。

|プロパティ名|内容|
|---|---|
|``addSample.sampleArtSheet``|テンプレートファイル内のサンプルアートがあるシートを指定して下さい。そのシートの画像を``-s``でコピーします。|
|``addSample.offsetL``|サンプル画像を貼り付ける際の左余白（列数）です。|
|``addSample.offsetT``|サンプル画像を貼り付ける際の上余白（行数）です。|
|``addSample.interval``|サンプル画像を貼り付ける際の貼り付け間隔（行数）です。|
|``addSample.endOfColumn``|シートの最大列数を指定します。これは空行の判別に使用します。|
|``addSample.endOfRow``|シートの最大行数を指定します。これ以降、サンプル画像は貼り付けられません。|
|``addSample.endBlankLine``|指定した行数以上空行が続いた場合、サンプル画像を貼り付けるのをやめます。|

# 留意事項

1. ``-sa``オプション実行中は、OSのクリップボード機能が使用不可になります。クリップボードに大事な情報が残っていないことを確認してから実行して下さい。

1. 処理対象のファイル名に``< > ? [ ] : | *``が含まれている場合、正常に処理を行うことが出来ません。

1. 処理中に``Ctrl + C``等で強制終了した場合、Excelのプロセスが残っている為、手動でタスクマネージャ等からKillして下さい。

# 既知の不具合

大量のファイルを処理している場合、``-sa``実行中にサンプル画像貼り付け中に失敗することがある。エラーコード:``「0x800A03EC」``

# License
MIT
