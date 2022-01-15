---
title: "[教えて！VBA] 第12回 各ワークシートのデータを集約して出力するにはどうすればいいの？？"
author: dede-20191130
date: 2022-01-15T14:20:46+09:00
slug: vba-question-012-aggregate-wss-content
draft: false
toc: true
featured: false
tags: ["VBA","Excel"]
categories: ["プログラミング"]
vba_taxo: vbaq
archives:
    - 2022
    - 2022-01
---


![](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1642259173/learnerBlog/vba-question-012-aggregate-wss-content/aggregate-wss-content1_uyhlmz.png)

## この記事について

{{< box-with-title title="かんたんな概要と結論" >}} 
    データが記載された各ワークシートにアクセスし、データを集約する方法について、<br/>
    サンプルを例示し、そのコードをベースにまとめました。<br/><br/>
    <code>WorkSheets</code>コレクションの取り扱いと、集約後の加工についても少し触れています。
{{< /box-with-title >}}

こんにちは、dedeです。

この記事では、  
VBAマクロに関する質問のうち、  
皆が疑問に思っているトピックについて解説いたします。

今回は、  
{{< colored-span color="#fb9700" >}}
ブックの各ワークシートのデータを集約して、他のワークシートやテキストファイルなどに出力する方法
{{< /colored-span >}}  
を解説いたします。

※この記事は、Office VBAマクロのうち  
Excel VBAマクロに関するトピックです。

レベル：<b>初級者向け</b>

## 環境

以下は、  
Office 2016のExcel環境で検証済みです。  

※2022/1時点の最新バージョンのExcelでも内容は変わりません。

## ワークシートの情報にアクセス
### ワークシートオブジェクトとは

各ワークシートの情報にアクセスするには、  
ワークシートオブジェクトというオブジェクトを参照します。

{{% notice tip VBAにおけるオブジェクト %}}
VBAはVisual Basic 6.0の仕様に準拠しており、  
オブジェクト指向のプログラミング言語です。

データは基本的にオブジェクト（データとそれに紐付く動作を一つにまとめたもの）として扱われます。  
オブジェクトはクラスによって生成され、クラスはオブジェクトの雛形とみなすことができます。

詳しくは、[こちらの記事]({{< relref "posts/2020/11/Excel-Class-Sample-01/index.md" >}})で紹介しました。
{{% /notice %}}

{{% notice info アクセスするコードサンプル %}}
ワークシートオブジェクトは、  
_ActiveSheet_（表示中のワークシート）プロパティや、  
_WorkSheets_ コレクションを呼び出すことなどで参照できます。

_WorkSheets_ コレクションについては下記で詳細に見ていきます。

```vb
Sub ワークシートオブジェクトの情報にアクセスする()
    Debug.Print ThisWorkbook.Worksheets(1).name '//ブックの一枚目のワークシートの名前を参照
    Debug.Print ThisWorkbook.Worksheets(1).Range("B2").Value '//ブックの一枚目のワークシートの、B2セルに書き込まれた値を参照
End Sub
```
{{% /notice %}}
### ワークシートコレクションとは

ワークシートコレクションとは、  
そのブックのワークシートがすべて格納されたコレクションを指します。

{{% notice tip "コレクションとは？" %}}
VBAにおけるコレクションとは、  
同じ種類のオブジェクトやデータを一定の規則に従って格納し、引き出すことができるようにしたオブジェクトの集まりで、  
それ自身もオブジェクトです。

「一定の規則」とは、  
_Collection.Item_(インデックス番号)や _Collection_ (インデックス番号)の形で〇〇番目のオブジェクトやデータを取得したり、  
_Collection.Count_ の形でコレクションに格納されたオブジェクトの数を算出したりする決まりを指します。    
さらに _For_ ループを用いて順番にオブジェクトのデータを処理できるなどの決まりも含まれます。

コレクションは、  
_New Collection_ の形で自分で定義できる方法もあれば、  
_WorkSheets_ のようにもともと定義されているものを即座に利用可能な場合も存在します。

```vb

Sub コレクション利用サンプル()
    Dim myColl As Collection '//自作のコレクションを宣言
    
    Set myColl = New Collection '//自作のコレクションをインスタンス化（準備）
    
    myColl.Add "りんご" '//自作のコレクションに文字列を格納
    myColl.Add "みかん"
    myColl.Add "ぶどう"
    
    Debug.Print myColl.Item(1) '//自作のコレクションから1番目のデータを参照
    Debug.Print myColl(2) '//自作のコレクションから2番目のデータを参照
    Debug.Print myColl.Count '//自作のコレクションのデータ数を算出
    
End Sub

```

{{% /notice %}}

ワークシートコレクションはExcelで既に定義されたオブジェクトで、コレクションの一種のため、  
上述したように`For`ループを用いて  
順番にブックのワークシートの情報を調べることができます。

```vb

Sub ワークシートコレクションをループするサンプル()
    Dim counter As Long
    Dim currentSheet As Worksheet '//ForEachループで使用するワークシートオブジェクト変数
    
    '//ループ処理によって各シートの名前を取得したい
    '////インデックス番号のカウンターを用いたForループの場合
    For counter = 1 To ThisWorkbook.Worksheets.Count
        Debug.Print ThisWorkbook.Worksheets(counter).name
    Next counter
    
    '////For Each機構を利用したループの場合
    For Each currentSheet In ThisWorkbook.Worksheets
        Debug.Print currentSheet.name
    Next currentSheet
    
End Sub

```

### ワークシートコレクションが属するブック

ワークシートコレクションを使用する際に注意すべきことは、  
調べたいワークシートコレクションがどのブックに属しているかということを明確に理解しておくべきということです。

下のコードでも示すように、  
`Worksheets`でどのブックのワークシートを取り扱うかということは、  
`Worksheets`のドットの前のブックオブジェクトがどのようなブックを指しているかということで決まります。

```vb

Sub ワークシートコレクションが属するブック()
    Dim targetWSColl As Sheets
    Dim newBook As Workbook
    
    Set targetWSColl = ThisWorkbook.Worksheets '// 1. マクロを記述しているブックのワークシートコレクションを変数に入れる
    
    Set targetWSColl = ActiveWorkbook.Worksheets '// 2. 現在開いているブックのワークシートコレクションを変数に入れる
    
    Set newBook = Workbooks.Add '// 新しくブックを作成
    Set targetWSColl = newBook.Worksheets '// 3. 新しく作成したブックのワークシートコレクションを変数に入れる
    newBook.Close False
End Sub

```

現在開いているブックのシートを取り扱いたいのに、  
`ThisWorkbook.Worksheets`を参照したりすることがないように気をつける必要があります。

### 各シートから取得したい範囲のデータを集約するサンプル

各シートのA1~A3セルの内容をメッセージ本文として集約するサンプルです。

```vb

Sub 各シートから取得したい範囲のデータを集約するサンプル()
    Dim currentSheet As Worksheet
    Dim aggregatedDataText As String
     '// 全シートについてセルの内容を調査
    For Each currentSheet In ThisWorkbook.Worksheets
        '// A1セルの内容取得
        aggregatedDataText = aggregatedDataText & currentSheet.Range("A1").Value & ";"
        '// A2セルの内容取得
        aggregatedDataText = aggregatedDataText & currentSheet.Range("A2").Value & ";"
        '// A3セルの内容取得
        aggregatedDataText = aggregatedDataText & currentSheet.Range("A3").Value
        aggregatedDataText = aggregatedDataText & vbLf
    Next currentSheet
    
    aggregatedDataText = "各シートから集めたデータはこちら" & vbLf & vbLf & aggregatedDataText
    
    '// 取得した内容をメッセージとして表示
    MsgBox aggregatedDataText, vbOKOnly, "各シートから取得したい範囲のデータを集約するサンプル"
End Sub

```


## データを出力する

次は、取得対象のデータを集約し、出力する処理について記します。

### 前提：シートに存在するデータの様式がそろっていること

まず前提として、  
シート状の、取得対象のデータの様式が揃っていなければなりません。

具体的には、  
取得したいセル範囲、セル結合しているか否か、各列の内容とその順番が統一されている必要があります。  
セル範囲がワークシートごとにまちまちだったり、項目の順番があるシートだけ逆順になっていたりするとコードが複雑化してしまいます。

今回は、  
２列のセル範囲を持つ売上金データをまとめてみます。

![売上金データの様式](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1642259173/learnerBlog/vba-question-012-aggregate-wss-content/%E3%82%B9%E3%82%AF%E3%83%AA%E3%83%BC%E3%83%B3%E3%82%B7%E3%83%A7%E3%83%83%E3%83%88_2022-01-15_180730_zuywba.png)

シートは、都道府県ごとのシートが複数と、  
例外となるシートも存在することとします。  
対象となるシートの名前は「〇〇データ」であるすべてのシートとします。

![シート名の様式](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1642259173/learnerBlog/vba-question-012-aggregate-wss-content/%E3%82%B9%E3%82%AF%E3%83%AA%E3%83%BC%E3%83%B3%E3%82%B7%E3%83%A7%E3%83%83%E3%83%88_2022-01-15_180747_qvnjvl.png)

なお、デモに用いたファイルは  
[こちら](https://github.com/dede-20191130/My_VBA_Tools/tree/master/Public/2022/01/aggregate-wss-content)からダウンロードできます。


### シートに出力する場合

まず、データを集めるシートと同じブックに、  
まとめ用シートを作成することを考えます。

![まとめ用シートの作成の流れ](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1642259173/learnerBlog/vba-question-012-aggregate-wss-content/aggregate-wss-content2_bevmse.png)

データの集約と出力は、  
コピーペーストではなく、値の取得・出力で行います

理由は、コピーペーストの場合、コメントや条件付き書式など、  
不必要なデータも伴ってしまう可能性があるからです。

```vb

Sub シートに出力する場合01_月単位でまとめないバージョン()
    Dim currentSheet As Worksheet '// ループ用
    Dim matomeSheet As Worksheet '// 作成したまとめシート
    Dim sheetDataDic As Object '//シートのデータ格納用のDictionary
    Dim currentSheetName As Variant '// ループ用
    Dim rowNumCursor As Long '// 操作する行番号を示すカーソル
    
    '// シートのデータ格納用のDictionaryを設定
    Set sheetDataDic = CreateObject("Scripting.Dictionary")
    
    '// 全シートごとにセルの内容を調査
    For Each currentSheet In ThisWorkbook.Worksheets
        '// 「〇〇データ」の名前のシートのみ調査
        If InStr(currentSheet.Name, "データ") > 0 Then
            sheetDataDic.Add currentSheet.Name, currentSheet.Range("B8:C19").Value
        End If
        '// まとめシートがすでにある場合は削除
        If currentSheet.Name = "まとめ" Then
            Application.DisplayAlerts = False
            currentSheet.Delete
            Application.DisplayAlerts = True
        End If
    Next currentSheet
    
    '// まとめシートを作成
    Set matomeSheet = ThisWorkbook.Worksheets.Add
    matomeSheet.Name = "まとめ"
    
    '// まとめシートに各シートの売上金などのデータを貼り付け
    rowNumCursor = 3
    '//// ヘッダー文字列を設定
    matomeSheet.Cells(2, 2).Value = "都道府県"
    matomeSheet.Cells(2, 3).Value = "計上月"
    matomeSheet.Cells(2, 4).Value = "売上金"
    '//// 各シートごとにデータを処理
    For Each currentSheetName In sheetDataDic.Keys
        '// 都道府県名を書き込み
        matomeSheet.Range( _
        matomeSheet.Cells(rowNumCursor, 2), _
        matomeSheet.Cells(rowNumCursor + 11, 2) _
        ).Value = currentSheetName
        '// 売上金などデータを書き込み
        matomeSheet.Range( _
        matomeSheet.Cells(rowNumCursor, 3), _
        matomeSheet.Cells(rowNumCursor + 11, 4) _
        ).Value = sheetDataDic(currentSheetName)
        
        '// 次の都道府県へ
        rowNumCursor = rowNumCursor + 12
    Next currentSheetName
    
End Sub

```

実行すると、  
「まとめ」シートが作成されます。

![「まとめ」シート](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1642259173/learnerBlog/vba-question-012-aggregate-wss-content/%E3%82%B9%E3%82%AF%E3%83%AA%E3%83%BC%E3%83%B3%E3%82%B7%E3%83%A7%E3%83%83%E3%83%88_2022-01-15_212138_hzftix.png)

すべての都道府県シートにおいて`B8:C19`の範囲でデータが有るため、  
コードが簡潔に書けています。

### シートに出力する場合（並べ替え処理を追加）

次に、上の関数に、  
計上月順に並べ替えを行う処理を追加します。

```vb

Sub シートに出力する場合02_月単位でまとめるバージョン()
    Dim currentSheet As Worksheet                '// ループ用
    Dim matomeSheet As Worksheet                 '// 作成したまとめシート
    Dim sheetDataDic As Object                   '//シートのデータ格納用のDictionary
    Dim currentSheetName As Variant              '// ループ用
    Dim rowNumCursor As Long                     '// 操作する行番号を示すカーソル
    
    '// シートのデータ格納用のDictionaryを設定
    Set sheetDataDic = CreateObject("Scripting.Dictionary")
    
    '// 全シートごとにセルの内容を調査
    For Each currentSheet In ThisWorkbook.Worksheets
        '// 「〇〇データ」の名前のシートのみ調査
        If InStr(currentSheet.Name, "データ") > 0 Then
            sheetDataDic.Add currentSheet.Name, currentSheet.Range("B8:C19").Value
        End If
        '// まとめシートがすでにある場合は削除
        If currentSheet.Name = "まとめ" Then
            Application.DisplayAlerts = False
            currentSheet.Delete
            Application.DisplayAlerts = True
        End If
    Next currentSheet
    
    '// まとめシートを作成
    Set matomeSheet = ThisWorkbook.Worksheets.Add
    matomeSheet.Name = "まとめ"
    
    '// まとめシートに各シートの売上金などのデータを貼り付け
    rowNumCursor = 3
    '//// ヘッダー文字列を設定
    matomeSheet.Cells(2, 2).Value = "都道府県"
    matomeSheet.Cells(2, 3).Value = "計上月"
    matomeSheet.Cells(2, 4).Value = "売上金"
    '//// 各シートごとにデータを処理
    For Each currentSheetName In sheetDataDic.Keys
        '// 都道府県名を書き込み
        matomeSheet.Range( _
        matomeSheet.Cells(rowNumCursor, 2), _
        matomeSheet.Cells(rowNumCursor + 11, 2) _
        ).Value = currentSheetName
        '// 売上金などデータを書き込み
        matomeSheet.Range( _
        matomeSheet.Cells(rowNumCursor, 3), _
        matomeSheet.Cells(rowNumCursor + 11, 4) _
        ).Value = sheetDataDic(currentSheetName)
        
        '// 次の都道府県へ
        rowNumCursor = rowNumCursor + 12
    Next currentSheetName
    
    '// 計上月と都道府県の列を入れ替える
    matomeSheet.Columns(3).Cut
    matomeSheet.Columns(2).Insert Shift:=xlToRight
    
    '// 計上月順に並び替えを行う
    With matomeSheet.Sort
        .SortFields.Clear
        .SortFields.Add2 Key:=matomeSheet.Range( _
        matomeSheet.Cells(3, 2), _
        matomeSheet.Cells(rowNumCursor - 1, 2) _
        ) _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange matomeSheet.Range( _
        matomeSheet.Cells(3, 2), _
        matomeSheet.Cells(rowNumCursor - 1, 4) _
        )
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
End Sub

```

実行すると、  
「まとめ」シートが作成され、  
なおかつデータが月の若い順からソートされて表示されます。

![ソートされたまとめシート](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1642259174/learnerBlog/vba-question-012-aggregate-wss-content/%E3%82%B9%E3%82%AF%E3%83%AA%E3%83%BC%E3%83%B3%E3%82%B7%E3%83%A7%E3%83%83%E3%83%88_2022-01-15_214619_unhfip.png)

このように、  
集約して別シートとしてデータをまとめれば、  
並び替えや分析も容易になります。

### シートに出力する場合（書式も合わせる）

今度は、  
データのテーブルの書式も都道府県シートの元データテーブルに  
寄せてみます。

```vb

Sub シートに出力する場合03_書式も合わせるバージョン()
    Dim currentSheet As Worksheet                '// ループ用
    Dim matomeSheet As Worksheet                 '// 作成したまとめシート
    Dim sheetDataDic As Object                   '//シートのデータ格納用のDictionary
    Dim currentSheetName As Variant              '// ループ用
    Dim rowNumCursor As Long                     '// 操作する行番号を示すカーソル
    
    '// シートのデータ格納用のDictionaryを設定
    Set sheetDataDic = CreateObject("Scripting.Dictionary")
    
    '// 全シートごとにセルの内容を調査
    For Each currentSheet In ThisWorkbook.Worksheets
        '// 「〇〇データ」の名前のシートのみ調査
        If InStr(currentSheet.Name, "データ") > 0 Then
            sheetDataDic.Add currentSheet.Name, currentSheet.Range("B8:C19").Value
        End If
        '// まとめシートがすでにある場合は削除
        If currentSheet.Name = "まとめ" Then
            Application.DisplayAlerts = False
            currentSheet.Delete
            Application.DisplayAlerts = True
        End If
    Next currentSheet
    
    '// まとめシートを作成
    Set matomeSheet = ThisWorkbook.Worksheets.Add
    matomeSheet.Name = "まとめ"
    
    '// まとめシートに各シートの売上金などのデータを貼り付け
    rowNumCursor = 3
    '//// ヘッダー文字列を設定
    matomeSheet.Cells(2, 2).Value = "都道府県"
    matomeSheet.Cells(2, 3).Value = "計上月"
    matomeSheet.Cells(2, 4).Value = "売上金"
    '//// 各シートごとにデータを処理
    For Each currentSheetName In sheetDataDic.Keys
        '// 都道府県名を書き込み
        matomeSheet.Range( _
        matomeSheet.Cells(rowNumCursor, 2), _
        matomeSheet.Cells(rowNumCursor + 11, 2) _
        ).Value = currentSheetName
        '// 売上金などデータを書き込み
        matomeSheet.Range( _
        matomeSheet.Cells(rowNumCursor, 3), _
        matomeSheet.Cells(rowNumCursor + 11, 4) _
        ).Value = sheetDataDic(currentSheetName)
        
        '// 次の都道府県へ
        rowNumCursor = rowNumCursor + 12
    Next currentSheetName
    
    '// 計上月と都道府県の列を入れ替える
    matomeSheet.Columns(3).Cut
    matomeSheet.Columns(2).Insert Shift:=xlToRight
    
    '// 計上月順に並び替えを行う
    With matomeSheet.Sort
        .SortFields.Clear
        .SortFields.Add2 Key:=matomeSheet.Range( _
                               matomeSheet.Cells(3, 2), _
                         matomeSheet.Cells(rowNumCursor - 1, 2) _
                         ) _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange matomeSheet.Range( _
                  matomeSheet.Cells(3, 2), _
                  matomeSheet.Cells(rowNumCursor - 1, 4) _
                  )
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    '// 書式を元のデータテーブルに合わせる
    '//// 罫線
    matomeSheet.Range( _
        matomeSheet.Cells(2, 2), _
        matomeSheet.Cells(rowNumCursor - 1, 4) _
        ).Borders.LineStyle = xlContinuous
    '//// ヘッダー背景色、中央揃え
    With matomeSheet.Range( _
         matomeSheet.Cells(2, 2), _
         matomeSheet.Cells(2, 4) _
         )
        .Interior.Color = 15917529
        .HorizontalAlignment = xlCenter
    End With
End Sub

```

実行すると、  
「まとめ」シートが作成され、ソートされ、  
なおかつヘッダーの背景色や罫線が元テーブル仕様に変更されます。

![書式が合わせられたまとめシート](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1642259174/learnerBlog/vba-question-012-aggregate-wss-content/%E3%82%B9%E3%82%AF%E3%83%AA%E3%83%BC%E3%83%B3%E3%82%B7%E3%83%A7%E3%83%83%E3%83%88_2022-01-15_233308_s8ii4v.png)

### テキストファイルに出力する場合

シートに出力せずにテキストファイルに出力する場合でも、  
いったん作業用シートとして一つのシートにまとめ、  
並び替えなどを行った後にテキストに出力すると便利な場合があるでしょう。

![テキストに出力する流れ](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1642259173/learnerBlog/vba-question-012-aggregate-wss-content/aggregate-wss-content3_kglhkg.png)

コードは以下のようになります。

```vb

Sub テキストファイル_CSV_に出力する場合()
    Dim currentSheet As Worksheet                '// ループ用
    Dim matomeSheet As Worksheet                 '// 作成したまとめシート
    Dim sheetDataDic As Object                   '//シートのデータ格納用のDictionary
    Dim currentSheetName As Variant              '// ループ用
    Dim rowNumCursor As Long                     '// 操作する行番号を示すカーソル
    
    '// シートのデータ格納用のDictionaryを設定
    Set sheetDataDic = CreateObject("Scripting.Dictionary")
    
    '// 全シートごとにセルの内容を調査
    For Each currentSheet In ThisWorkbook.Worksheets
        '// 「〇〇データ」の名前のシートのみ調査
        If InStr(currentSheet.Name, "データ") > 0 Then
            sheetDataDic.Add currentSheet.Name, currentSheet.Range("B8:C19").Value
        End If
        '// まとめシートがすでにある場合は削除
        If currentSheet.Name = "まとめ" Then
            Application.DisplayAlerts = False
            currentSheet.Delete
            Application.DisplayAlerts = True
        End If
    Next currentSheet
    
    '// まとめシートを作成
    Set matomeSheet = ThisWorkbook.Worksheets.Add
    matomeSheet.Name = "まとめ"
    
    '// まとめシートに各シートの売上金などのデータを貼り付け
    rowNumCursor = 2
    '//// ヘッダー文字列を設定
    matomeSheet.Cells(1, 1).Value = "都道府県"
    matomeSheet.Cells(1, 2).Value = "計上月"
    matomeSheet.Cells(1, 3).Value = "売上金"
    '//// 各シートごとにデータを処理
    For Each currentSheetName In sheetDataDic.Keys
        '// 都道府県名を書き込み
        matomeSheet.Range( _
        matomeSheet.Cells(rowNumCursor, 1), _
        matomeSheet.Cells(rowNumCursor + 11, 1) _
        ).Value = currentSheetName
        '// 売上金などデータを書き込み
        matomeSheet.Range( _
        matomeSheet.Cells(rowNumCursor, 2), _
        matomeSheet.Cells(rowNumCursor + 11, 3) _
        ).Value = sheetDataDic(currentSheetName)
        
        '// 次の都道府県へ
        rowNumCursor = rowNumCursor + 12
    Next currentSheetName
    
    '// 計上月と都道府県の列を入れ替える
    matomeSheet.Columns(2).Cut
    matomeSheet.Columns(1).Insert Shift:=xlToRight
    
    '// 計上月順に並び替えを行う
    With matomeSheet.Sort
        .SortFields.Clear
        .SortFields.Add2 Key:=matomeSheet.Range( _
                               matomeSheet.Cells(3, 1), _
                         matomeSheet.Cells(rowNumCursor - 1, 1) _
                         ) _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange matomeSheet.Range( _
                  matomeSheet.Cells(3, 1), _
                  matomeSheet.Cells(rowNumCursor - 1, 3) _
                  )
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    '// まとめシートを別ブックとして取り出す
    matomeSheet.Move
    
    '// まとめシートをCSV形式のテキストファイルとしてブックと同じフォルダに保存
    ActiveWorkbook.SaveAs Filename:=ThisWorkbook.Path & "\" & "まとめ.csv" _
        , FileFormat:=xlCSVUTF8, CreateBackup:=False
    ActiveWorkbook.Close
    
End Sub

```

{{% notice info Moveは必須です %}}
_matomeSheet.Move_ は、  
CSVとして保存する前にかならず行うべき処理です。

これをせずにCSV形式でブックを保存すると、  
マクロが記載されたブックをCSVに変換することになり、不本意な結果に繋がる可能性があります。
{{% /notice %}}


## 終わりに

ワークシートのデータを集約し、  
まとめる方法についていくつかのバリエーションを見てきました。

より複雑なケースにも対応できるように、  
ワークシートに記載するデータの様式を整えておくと便利でしょう。
