---
title: "[教えて！VBA] 第2回 一つのセルの値が「東京都」であることを条件として別のアクションを行うにはどうすればいいの？？"
author: dede-20191130
date: 2021-12-14T11:53:35+09:00
slug: vba-question-002-chekc-single-cell
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

## この記事について

{{< box-with-title title="かんたんな概要と結論" >}} 
    Excelのシートのセルに入力されたデータを読み取るには、<br>
    RangeあるいはCellsを使用します。<br><br>
    セルのデータを読み取り処理を様々に書くことができれば、<br>
    VBAマクロの表現の幅は非常に広がります。
{{< /box-with-title >}}

こんにちは、dedeです。

この記事では、  
VBAマクロに関する質問のうち、  
皆が疑問に思っているトピックについて解説いたします。

今回は、
{{< colored-span color="#fb9700" >}}一つのセルの値が「東京都」であることをチェックする方法とその比較について{{< /colored-span >}}  
を解説いたします。

※この記事は、Office VBAマクロのうち  
Excel VBAマクロに関するトピックです。

レベル：<b>初級者向け</b>


## セルの値を調べたいのはどんな場合？

例えば、特定の値が入っているセルだけを対象として、  
その隣のセルに文字を入力したい場合を考えます。

下記の図では、  
都道府県が「東京都」と入力されたセルの隣のセルだけに「必要」と自動入力するマクロを想定しています。  

![想定する流れ](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1639785261/learnerBlog/vba-question-002-chekc-single-cell/vba-question-002-chekc-single-cell_g8kyeg.png)

他にも、その値の入ったセル自身の背景色を変更したり、  
値を取得してなにかの計算処理を加えて他のセルに反映するなど、  
{{< colored-span color="#fb9700" >}}Excel VBAマクロにおいては、セルの値を参照することはもっとも使われる基本的なテクニックです。{{< /colored-span >}}

## 方法
### セルの値を取得する

セルの値を取得するには、  
`Rangeプロパティ`を使用する、あるいは`Cellsプロパティ`を使用することで実現できます。

いずれにしても、`Rangeオブジェクト`（プロパティと同名なのでややこしい）という  
セル範囲の情報をもっているオブジェクトを取得する点では相違ありません。

※ここでいう「オブジェクト」とは、  
対象としているものの情報とそれを操作する関数をセットで持っているデータの一種と理解ください。  

例えばRangeオブジェクトはセル範囲のアドレスや、セルをコピーする操作関数を持っています。


#### Rangeを使う

次のようにRangeプロパティの引数にセルのアドレスを指定します。

```vb
Dim cellValue As String
cellValue = Range("B2").Value
```

このようにすると、B2セルに記載された内容を  
文字列として`cellValue`変数に格納します。

例えば、B2セルに「あいうえお」と入っていた場合、  
`cellValue`にも「あいうえお」と入ります。

#### Cellsを使う

次のようにRangeプロパティの引数にセルのアドレスを{{< colored-span color="#fb9700" >}}座標として{{< /colored-span >}}指定します。

```vb
Dim cellValue As String
cellValue = Cells(2, 2).Value
```

`Cells(2, 2)`というのはB2セルを指しています。

気づかれたかもしれませんが、  
`Cells`の引数の最初には対象セルの行番号が、その次には対象セルの列番号を数値に変換したものが入ります。

![Cellsの引数規則](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1639785261/learnerBlog/vba-question-002-chekc-single-cell/vba-question-002-chekc-single-cell2_j6zawz.png)


#### どちらを使えばいいの？

`Range`でも`Cells`でも指定したセルの内容を取得できるとわかりましたが、  
どのように使い分けるのか？？

それは、  
`Range`には一つのセルだけでなく複数のセルの情報をオブジェクトとして取得し、それらを操作できるという利点があるが、  
`Cells`は常に単一のセルを取り扱うプロパティだということに焦点を当てれば見えてきます。

例えば、  
次のコードでは、
A1セルからC3セルまでのセルの背景色を一括で赤色に変更します。

```vb
Range("A1:C3").Interior.Color = 255
```

そのため、  
このように一括で複数セルを取り扱いたいときに`Range`を使用します。


一方、`Cells`を使った場合、  
プログラムコードを見たときに単一のセルを取り扱っていることがすぐにわかるので、  
「あっ、このコードのここの部分では一つのセルについて操作しているんだな」とわかりやすくなります。

上記の理由から、  
複数セルを取り扱いたい場合以外では、`Cells`を使用することを勧めます。


### セルの値によって分岐させる

次に、  
「IFステートメント」というVBAの構文を使って  
取得したセルの値に応じて処理を分岐させます。

```vb
If 評価式 Then
    ここは評価式がYESであった場合にのみ実行されます。    
Else
    ここは評価式がNOであった場合にのみ実行されます。    
End If
```

「評価式」とは、  
内容がYES/NO（VBAの用語ではTRUE/FALSE）として表現できる式が入ります。  

例えば「1 + 2 = 3」「東京は日本の首都」「滋賀は東北地方の県である」はすべて評価式です。

上記のIFステートメントを使って、  
取得したセルの値が「東京都」であった場合にのみ隣のセルの値に「必要」と入力するコードは  
次のようになります。

```vb
If cellValue = "東京都" Then
    Cells(隣のセルの行番号, 隣のセルの列番号).Value = "必要"
End If
```

なお、ここでは評価式がNOであった場合には  
何もしないので、  
`Else`ステートメント部分は使用していません。

## デモ

### Cellsで書かれたデモ

都道府県が「東京都」と入力されたセルの隣のセルだけに「必要」と自動入力するマクロです。

```vb
Sub 必要の入力()
    
    If Cells(3, 2).Value = "東京都" Then
        Cells(3, 3).Value = "必要"
    End If
    If Cells(4, 2).Value = "東京都" Then
        Cells(4, 3).Value = "必要"
    End If
    If Cells(5, 2).Value = "東京都" Then
        Cells(5, 3).Value = "必要"
    End If
    If Cells(6, 2).Value = "東京都" Then
        Cells(6, 3).Value = "必要"
    End If
    
End Sub
```

これを実行すると、  
C3セルとC6セルにのみ「必要」と入力されます。

![実行前](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1639785260/learnerBlog/vba-question-002-chekc-single-cell/01_x4vft6.png)

![実行後](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1639785260/learnerBlog/vba-question-002-chekc-single-cell/02_jwnums.png)


### Range形式に書き直すと……

Range形式に書き直すと、次のように書けます。  
もともとExcel関数をよく使っていたユーザにはこちらのほうが馴染み深いかもしれません。

```vb
Sub 必要の入力()
    
    If Range("B3").Value = "東京都" Then
        Range("C3").Value = "必要"
    End If
    If Range("B4").Value = "東京都" Then
        Range("C4").Value = "必要"
    End If
    If Range("B5").Value = "東京都" Then
        Range("C5").Value = "必要"
    End If
    If Range("B6").Value = "東京都" Then
        Range("C6").Value = "必要"
    End If
    
End Sub
```

## 終わりに

セルからデータを取得して、  
それに応じた処理を行うことは、  
Excel VBAマクロでもっともよく使われるロジックです。

これをうまく使えるようになれば、  
マクロの自動化の幅が非常に広がるでしょう。

ここには書ききれませんでしたが、  
`Range`や`Cells`の使用法には他にもバラエティがあるので、  
Microsoftの公式を見てみると面白いかもしれません。