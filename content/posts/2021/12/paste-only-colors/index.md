---
title: "[Excel VBA] セルの文字色・背景色だけをコピーして貼り付けるマクロを作成しました"
author: dede-20191130
date: 2021-12-29T10:24:02+09:00
slug: paste-only-colors
draft: false 
toc: true
featured: false
tags: ["VBA","Excel"]
categories: ["プログラミング"]
archives:
    - 2021
    - 2021-12
---

![](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1640748357/learnerBlog/paste-only-colors/paste-only-colors_osfx92.png)

## この記事について

{{< box-with-title title="かんたんな概要と結論" >}} 
    かんたんな関数と変数の組み合わせで、<br>
    文字色・背景色だけをコピペするマクロは作成できます。
{{< /box-with-title >}}

こんにちは、dedeです。

この記事では、  
セルや図形の持っている書式に関する各設定のうち、  
色（背景色・文字色）だけを他のセル・図形にペーストするマクロについてご説明します。

## 作成環境

- Microsoft Office Excel 2016

## やりたいこと

Excelで事務作業等をしているときに、  
セルの色だけを他のセルにコピーペーストさせたい時があるかと思います。

普通に「書式の貼り付け」を行うと、  
色以外の情報（文字サイズ、罫線など）も反映させてしまうため、それが不都合な場合もあります。

![書式の貼り付け](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1640748356/learnerBlog/paste-only-colors/%E3%82%B9%E3%82%AF%E3%83%AA%E3%83%BC%E3%83%B3%E3%82%B7%E3%83%A7%E3%83%83%E3%83%88_2021-12-29_113515_ka7lae.png)

また、結合セルから他の結合セルにコピペする場合を考えると、  
そもそも書式貼り付けをすると結合が崩れてしまうため、  
使えません。

![結合セルから他の結合セルに書式をコピペできない](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1640748357/learnerBlog/paste-only-colors/paste-only-colors2_xbeqmn.png)

そこで、VBAの出番です。  

※なお、上記ではセルについて説明しましたが、  
これから説明するマクロは、
セルだけではなく図形についても同様に色のコピペが可能です。

## コード
### ABOUT

全体のコードは役割ごとに次のように分割されます。  
- モジュール変数
- 色をコピーする関数
- 他のセル・図形に貼り付けする関数


### モジュール変数

作成したモジュールの先頭で、  
文字色格納用変数`fontColor`と背景色格納用変数`backgroundColor`を宣言します。

VBAでは色を16進数カラーコードとして取り扱うので、  
変数の型は`LONG型`です。

```vb{hl_lines=[3,4]}
Option Explicit

Private fontColor As Long
Private backgroundColor As Long

Sub コピーする関数()
'do something...
End Sub

Sub 貼り付けする関数()
'do something...
End Sub

```

### コピーする関数

```vb
Sub コピーする関数()
    fontColor = Selection.Font.Color
    backgroundColor = Selection.Interior.Color
End Sub

```

### 貼り付けする関数

```vb
Sub 貼り付けする関数()
    Selection.Font.Color = fontColor
    Selection.Interior.Color = backgroundColor
End Sub


```

### 使用方法

まずセルを選択し、  
コピー関数を実行します。  

次に色をペーストしたいセル（一つのセルでもセル範囲でも）を選択し、  
貼り付け関数を実行します。

実行の様子は「デモ」セクションでご説明します。

### なぜこれでうまくいくのか？

上記で見る通り、  
いたってシンプルな関数の組み合わせです。

コツは、  
「モジュール変数に色のコードを格納すること」です。

VBAの変数にはライフサイクル（使用できる期限）が存在し、  
関数の中で宣言した変数は関数が終了すると使えなくなりますが、  
モジュール変数は関数終了後もメモリに格納され続けるため、  
マクロが入ったExcelブックを終了するまでは使い続けることができます。

よって、コピー関数を実行した後に入れられた色のコードを使い回すことができるのです。

## デモ
### セルからセルにコピー
下記の図表②において、  
図表①のように、3000円以上の行を黄色く塗りたいということを想定します。

![デモ用図表](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1640748356/learnerBlog/paste-only-colors/%E3%82%B9%E3%82%AF%E3%83%AA%E3%83%BC%E3%83%B3%E3%82%B7%E3%83%A7%E3%83%83%E3%83%88_2021-12-29_121218_paqas0.png)

①まず、コピーしたい色の情報を持ったセルを選択後、  
「コピーする関数」を実行します。

![セルを選択](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1640748356/learnerBlog/paste-only-colors/%E3%82%B9%E3%82%AF%E3%83%AA%E3%83%BC%E3%83%B3%E3%82%B7%E3%83%A7%E3%83%83%E3%83%88_2021-12-29_120954_i27co0.png)

②その後、  
ペーストしたいセル範囲を選択し、  
「貼り付けする関数」を実行します。

![貼り付けの実行](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1640748356/learnerBlog/paste-only-colors/%E3%82%B9%E3%82%AF%E3%83%AA%E3%83%BC%E3%83%B3%E3%82%B7%E3%83%A7%E3%83%83%E3%83%88_2021-12-29_121422_zzzxj1.png)

色の情報が正しくペーストできました。

### 図形から図形にコピー
図形の間の色のやり取りも可能です。

まず、下の図の「あいうえお」の図形を選択し、  
コピー関数実行します。

![コピー元図形の選択](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1640748356/learnerBlog/paste-only-colors/%E3%82%B9%E3%82%AF%E3%83%AA%E3%83%BC%E3%83%B3%E3%82%B7%E3%83%A7%E3%83%83%E3%83%88_2021-12-29_121702_tbhxxc.png)

次に、色を反映させたい「かきくけこ」「さしすせそ」の図形を選択し、  
貼り付け関数を実行します。

![貼り付けの実行](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1640748356/learnerBlog/paste-only-colors/%E3%82%B9%E3%82%AF%E3%83%AA%E3%83%BC%E3%83%B3%E3%82%B7%E3%83%A7%E3%83%83%E3%83%88_2021-12-29_121729_nxeeuf.png)

図形においても色のコピペをすることができました。

### 図形からセルにコピー

最後に、図形からセルに色を反映させることもできることをデモします。  
（割愛しましたが、反対に、セルから図形に色を反映させることもできます。）

反映前：  
![反映前](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1640748356/learnerBlog/paste-only-colors/%E3%82%B9%E3%82%AF%E3%83%AA%E3%83%BC%E3%83%B3%E3%82%B7%E3%83%A7%E3%83%83%E3%83%88_2021-12-29_122055_wz3kjb.png)

反映後：  
![反映後](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1640748356/learnerBlog/paste-only-colors/%E3%82%B9%E3%82%AF%E3%83%AA%E3%83%BC%E3%83%B3%E3%82%B7%E3%83%A7%E3%83%83%E3%83%88_2021-12-29_122117_ntewzk.png)

## 終わりに

色だけを色々操作したいケースは、  
特に結合セルが多用された図表で色を統一させたい場合に  
よくあるかと思います。

ただ、よく言われていることですが、  
結合セルの多用自体がアンチパターン（あまり好ましくない処理）であることも考慮するべきではあるでしょう。

本当にこのセルの構造が最適か、ということを考慮した上で使用すると、  
より効果的かとおもわれます。
