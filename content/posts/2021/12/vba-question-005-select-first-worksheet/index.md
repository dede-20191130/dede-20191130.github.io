---
title: "[教えて！VBA] 第5回 ブックの最初のシートの最初のセルを選択した状態にするにはどうすればいいの？？"
author: dede-20191130
date: 2021-12-31T10:32:02+09:00
slug: vba-question-005-select-first-worksheet
draft: false
toc: true
featured: false
tags: ["VBA","Excel"]
categories: ["プログラミング"]
vba_taxo: vbaq
archives:
    - 2021
    - 2021-12
---

![](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1640920597/learnerBlog/vba-question-005-select-first-worksheet/select-first-worksheet_zrvw3w.png)

## この記事について

{{< box-with-title title="かんたんな概要と結論" >}} 
    各シートはSheetsコレクションに含まれ、<br>
    インデックスの最初の数をコレクションで指定することで最初のシートを参照できます。<br>
    非表示シートの存在には注意が必要です。
{{< /box-with-title >}}

こんにちは、dedeです。

この記事では、  
VBAマクロに関する質問のうち、  
皆が疑問に思っているトピックについて解説いたします。

今回は、
{{< colored-span color="#fb9700" >}}ブックの最初のシートの最初のセルを選択した状態にする方法{{< /colored-span >}}  
を解説いたします。

※この記事は、Office VBAマクロのうち  
Excel VBAマクロに関するトピックです。

レベル：<b>初級者向け</b>


## やりたいこと

Excelで作業をしていると、  
トピックごとに複数シートに分けて記載して内容を管理するというのは  
よく行われるかと思われます。

![複数シートで管理](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1640920597/learnerBlog/vba-question-005-select-first-worksheet/select-first-worksheet2_gfdkbu.png)

シートを他の人に渡したりするときに、  
ブックの最初のシートを選択した状態にしておきたい場合があるかと思います。  

下の図では、マクロで「表紙」シートを開いた状態にする流れです。

![最初のシートを表示するマクロ](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1640920597/learnerBlog/vba-question-005-select-first-worksheet/select-first-worksheet3_ekxzrb.png)

この記事ではその処理の流れを説明します。

## 処理の流れ

処理の流れは、大きく分けて
1. ブックの最初のシートを選択する処理
2. シートの最初のセルを選択する処理
の順です。

### ブックの最初のシートを選択する処理

シートを選択するためには、  
シートが格納されているものの実態について理解するところから始まる必要があります。

※※ここでいう「シート」とはいわゆるワークシート（通常使うセルが敷き詰められたシート）で、  
グラフシートやマクロシートと呼ばれるシートは含まれていないことを前提とします。  
これらのシートはあまり使われないですしね。

#### Sheetsコレクションとは？

わたしたちが作成したシートは、  
Sheetsコレクションという格納用変数に順次入れられていきます。

コレクションとは、  
VBAにおけるオブジェクト（データとそれを操作する関数をまとめたもの）の一種で、  
同じ構造のデータをまとめて格納するためのものです。

その点で、配列によく似ています。  
ただし、データの入れ方や取り出し方法などの点で配列とは区別されます。

![Sheetsコレクション](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1640920598/learnerBlog/vba-question-005-select-first-worksheet/select-first-worksheet4_hxsu1x.png)

現在表示されているブックのSheetsコレクションを取得するには、  
次のように書きます。

```vb
Set mySheets = ActiveWorkbook.Worksheets
Set mySheets = Worksheets '//省略も可

```



#### シートのインデックスとは？

シートのインデックスとは、  
シートの並び順を規定するもので、  
Sheetsコレクションで操作したいシートを参照する際に利用されます。

```vb
my_number = ActiveSheet.Index '//現在表示しているシートのインデックス番号が格納されます
name = Worksheets.Item(my_number).name '//現在表示しているシートを参照し、その名前を取得します。
name = Worksheets(my_number).name '//Itemプロパティは省略可。省略すると、配列に似た参照方法となります。
```

インデックスは1から始まりシート数を最大上限とします。  
10シート存在するブックなら、
最初のシートのインデックスは1、最後のシートのインデックスは10です。

インデックスの並びの決め方は、  
UIで見たときの左から右のシートタブの並びと同じです。  
例えば、下のようなブックであれば、  
表紙シートのインデックスは1、年間推移シートは2、総括シートは3です。

![ブックサンプル](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1640920597/learnerBlog/vba-question-005-select-first-worksheet/select-first-worksheet5_yrssti.png)

#### 選択方法

上記を踏まえ、  
ブックの最初のシートを選択する処理のコードは次のようになります。

```vb
Dim firstWorkSheet
Set firstWorkSheet = Worksheets(1)
firstWorkSheet.Select '//シートの選択

```

#### ただし非表示シートが有ると工夫が必要

今まで説明した方法は、  
じつは非表示シートが存在すると使用できません。

理由は、非表示シートではSelectメソッドを実行するとエラーが起こるからです。

![エラーの発生](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1640920596/learnerBlog/vba-question-005-select-first-worksheet/%E3%82%B9%E3%82%AF%E3%83%AA%E3%83%BC%E3%83%B3%E3%82%B7%E3%83%A7%E3%83%83%E3%83%88_2021-12-31_115927_atmj1c.png)

そのため、非表示ではないシートのうち  
最も番号の若いものを選択するという工夫が必要となります。  

```vb
'可視シートのうち最も番号の若いものを選択
For Each my_sheet In Worksheets
    If my_sheet.Visible = True Then
        my_sheet.Select
        Exit For
    End If
Next

```

マクロを利用するブックが、  
非表示シートを含んでいるかいないかについては  
注意する必要があるでしょう。

### シートの最初のセルを選択する処理

シートの最初のセルはA1セルです。

そのため、A1にカーソルを選択させれば目的達成されます。

```vb
ActiveSheet.Range("A1").Select
```
## 最終的なコード

```vb
Sub ブックの最初のシートの最初のセルを選択()
    Dim firstWorkSheet
    Set firstWorkSheet = Worksheets(1)
    firstWorkSheet.Select '//シートの選択
    
    firstWorkSheet.Range("A1").Select
End Sub

```
## デモ

次のように3シート存在するブックで、  
三枚目の総括シートを表示中です。

![マクロ実行前](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1640920596/learnerBlog/vba-question-005-select-first-worksheet/%E3%82%B9%E3%82%AF%E3%83%AA%E3%83%BC%E3%83%B3%E3%82%B7%E3%83%A7%E3%83%83%E3%83%88_2021-12-31_120748_iacajx.png)

VBAマクロを実行すると、  
一枚目の表紙シートが開かれ、A1セルに照準が合わされた状態になります。

![マクロ実行後](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1640920596/learnerBlog/vba-question-005-select-first-worksheet/%E3%82%B9%E3%82%AF%E3%83%AA%E3%83%BC%E3%83%B3%E3%82%B7%E3%83%A7%E3%83%83%E3%83%88_2021-12-31_120804_nyzckb.png)

## 応用：すべてのシートの最初のセルを選択した状態にするマクロ

一枚目シートだけではなく、  
すべてのシートのA1セルにカーソルが合った状態にするマクロは、  
次のようになります。

```vb
Sub すべてのシートの最初のセルを選択
    '変数
    Dim cntObj As Object                         'ループカウンタ
    
    
    'すべてのシートのカーソルを左上セル（A1セル）に合わせる
    For Each cntObj In Worksheets
        '非表示セルは飛ばす
        If cntObj.Visible = True Then
            cntObj.Select
            cntObj.Range("A1").Select
        End If
    Next
    
    '可視シートのうち最も番号の若いものを選択
    For Each cntObj In Worksheets
        If cntObj.Visible = True Then
            cntObj.Select
            Exit For
        End If
    Next

End Sub
```
## 終わりに
このマクロを通して、  
シートの取り扱い方、  
VBAにおけるコレクションという存在と使用方法、  
非表示シートの注意点について把握することができるでしょう。

これらの知識は他マクロにも応用できるでしょう。