---
title: "[教えて！VBA] 第4回 シート上の図形をコピーして他のセル上に貼り付けするにはどうすればいいの？？"
author: dede-20191130
date: 2021-12-30T13:16:26+09:00
slug: vba-question-004-copypaste-shapes
draft: false
toc: true
featured: false
tags: ["VBA","Access"]
categories: ["プログラミング"]
archives:
    - 2021
    - 2021-12
---

![](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1640865307/learnerBlog/vba-question-004-copypaste-shapes/copypaste-shapes_zbaepe.png)

## この記事について

{{< box-with-title title="かんたんな概要と結論" >}} 
    コピー元の図形を参照するためには少々下準備が必要となります。<br>
    ペーストには二種類の方法がありますが、どちらを使用しても目的を果たすことができます。
{{< /box-with-title >}}

こんにちは、dedeです。

この記事では、  
VBAマクロに関する質問のうち、  
皆が疑問に思っているトピックについて解説いたします。

今回は、
{{< colored-span color="#fb9700" >}}シート上の図形をコピーして他のセル上に貼り付けする方法{{< /colored-span >}}  
を解説いたします。

※この記事は、Office VBAマクロのうち  
Excel VBAマクロに関するトピックです。

レベル：<b>中級者向け</b>

## やりたいこと 

次のように、すでに存在する図形を使って、  
オレンジ色セルの位置に図形をコピーしたい、  
しかも「正確な位置」にコピーしたいというケースを考えます。

■■ 正確な位置に処理を行う、というのは典型的なVBAマクロの出番です ■■

![図形をコピーしたい](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1640865307/learnerBlog/vba-question-004-copypaste-shapes/copypaste-shapes2_fomoqm.png)

実装の方法としてはいくつか種類があるので、  
それぞれの方法と利点について記していきたいと思います。

## コード 
### 図形の選択 

コピー元の図形の選択方法は、  
あらかじめ名前を名付けたうえで選択する方法と、  
図形の置かれた位置のセル座標から選択する方法があります。

#### 方法① 名前を名付けて選択

新しく図形を作成すると、  
それぞれの図形には図形種類ごとに自動的に名前が割り振られます。

次の図のように笑顔マークを作成すると、  
リボンの左下にある「名前ボックス」に「スマイル１」と名前がつけられています。

![名前ボックスのデフォルト名](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1640865307/learnerBlog/vba-question-004-copypaste-shapes/copypaste-3shapes_o10c71.png)

{{< inner-article-div color="red" >}}しかし、この名前のままVBAで図形を参照することはできません。{{< /inner-article-div >}}

理由は、名前ボックスで表示される名前はローカライズされています（日本語）が、  
一方でVBAで使用される名前は英語であるためです。

```vb
Set myShape = ActiveSheet.Shapes("スマイル１")'//not working!
```

VBAで図形を参照するためには、  
作成した図形の一つ一つに新しく名前を名付ける必要があります。

名前の付け方は、  
図形を選択し、  
リボンの左下にある「名前ボックス」で名前を編集します。

![名前ボックスにおける名前の編集後](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1640865307/learnerBlog/vba-question-004-copypaste-shapes/copypaste-shapes4_xfwjyd.png)

笑顔マーク、星マーク、禁止記号をそれぞれ次のように変更します。

|図形種類|図形デフォルト名|図形変更後名前|
|-|-|-|
|笑顔マーク|スマイル 1|笑顔|
|星マーク|星: 5 pt 2|星|
|禁止記号|"禁止"マーク 3|禁止|

最後に、  
VBA上では次のように図形を参照できます。

```vb
Sub 名前を名付けて図形を選択()
    Dim shapeSmile As Shape
    Dim shapeStar As Shape
    Dim shapeNG As Shape
    
    Set shapeSmile = ActiveSheet.Shapes("笑顔")
    Set shapeStar = ActiveSheet.Shapes("星")
    Set shapeNG = ActiveSheet.Shapes("禁止")
    
    Debug.Print shapeSmile.TopLeftCell.Address '//笑顔マーク図形の座標位置
    Debug.Print shapeStar.TopLeftCell.Address '//星マーク図形の座標位置
    Debug.Print shapeNG.TopLeftCell.Address '//禁止マーク図形の座標位置
End Sub

```


#### 方法② 図形の置かれた位置のセル座標から選択

次のコードのように、  
あらかじめ対象となる図形の置かれたセル座標を把握した上で、  
現在のシートのすべての図形をループ処理で探索し、  
目的の座標に一致したものを変数に格納していきます。

```vb
Sub 図形の置かれた位置のセル座標から選択()
    Dim shapeSmile As shape
    Dim shapeStar As shape
    Dim shapeNG As shape
    
    Dim myShape As Object
    
    For Each myShape In ActiveSheet.Shapes
        Select Case myShape.TopLeftCell.Address
        '//笑顔マーク図形の座標に一致する場合
        Case "$G$4"
            Set shapeSmile = myShape
        '//星マーク図形の座標に一致する場合
        Case "$I$4"
            Set shapeStar = myShape
        '//禁止マーク図形の座標に一致する場合
        Case "$K$4"
            Set shapeNG = myShape
        End Select
        
    Next
    
    Debug.Print shapeSmile.TopLeftCell.Address '//笑顔マーク図形の座標位置
    Debug.Print shapeStar.TopLeftCell.Address '//星マーク図形の座標位置
    Debug.Print shapeNG.TopLeftCell.Address '//禁止マーク図形の座標位置
    
End Sub
```
#### どちらが良いのか？？

結論から言うと、  
処理速度重視なら方法①、動的に図形を生成するコードを書くのならば方法②に軍配が上がるのかなと思います。

②の方法ではループ処理でシート上のすべての図形を調べるため、  
どうしても時間はかかります（それでも図形の数が少なければ微々たるものですが）。

一方、コピー元の図形が次々と動的に生成されるような場合や、  
シートごとどこか他のブックからコピーして、その上の図形を操作したいような場合では、  
いちいち名前を変更するのも大変だし、名前被りの問題も考慮しなければならないため、  
方法②に優位性があるでしょう。

### コピーとペースト

図形をコピーして指定した場所にペーストする方法は、  
ペースト先のセルを選択した後にペーストを行う方法と、  
シートの任意の場所にペーストしてから図形を移動する方法の二種類が存在します。

#### 方法① ペースト先のセルを選択してペースト

次のようなコードです。

```vb
Sub ペースト先のセルを選択してペースト()
    Dim shapeSmile As shape
    Dim shapeStar As shape
    Dim shapeNG As shape
    
    Set shapeSmile = ActiveSheet.Shapes("笑顔")
    Set shapeStar = ActiveSheet.Shapes("星")
    Set shapeNG = ActiveSheet.Shapes("禁止")
    
    '//笑顔マークをペースト
    shapeSmile.Copy
    
    '////1. G5セル（ペーストしたいセルアドレス）を選択
    ActiveSheet.Range("G5").Select
    '////2. ペーストの実行
    ActiveSheet.Paste
    '////以下繰り返し
    ActiveSheet.Range("G6").Select
    ActiveSheet.Paste
    
    '//星マークをペースト
    shapeStar.Copy
    ActiveSheet.Range("I5").Select
    ActiveSheet.Paste
    ActiveSheet.Range("I6").Select
    ActiveSheet.Paste
    
    '//禁止マークをペースト
    shapeNG.Copy
    ActiveSheet.Range("K5").Select
    ActiveSheet.Paste
    ActiveSheet.Range("K6").Select
    ActiveSheet.Paste
    
End Sub
```

#### 方法② シートにペーストしてから図形を移動

次のようなコードです。

```vb
Sub シートにペーストしてから図形を移動()
    Dim shapeSmile As shape
    Dim shapeStar As shape
    Dim shapeNG As shape
    
    Set shapeSmile = ActiveSheet.Shapes("笑顔")
    Set shapeStar = ActiveSheet.Shapes("星")
    Set shapeNG = ActiveSheet.Shapes("禁止")
    
    '//笑顔マークをペースト
    shapeSmile.Copy
    '////1. ペーストの実行
    ActiveSheet.Paste
    '////2. ペーストした図形を指定位置（ペーストしたいセルアドレス位置）に移動
    With Selection
        .Top = ActiveSheet.Range("G5").Top
        .Left = ActiveSheet.Range("G5").Left
    End With
    '////以下繰り返し
    ActiveSheet.Paste
    With Selection
        .Top = ActiveSheet.Range("G6").Top
        .Left = ActiveSheet.Range("G6").Left
    End With
    
    '//星マークをペースト
    shapeStar.Copy
    ActiveSheet.Paste
    With Selection
        .Top = ActiveSheet.Range("I5").Top
        .Left = ActiveSheet.Range("I5").Left
    End With
    ActiveSheet.Paste
    With Selection
        .Top = ActiveSheet.Range("I6").Top
        .Left = ActiveSheet.Range("I6").Left
    End With
'
    '//禁止マークをペースト
    shapeNG.Copy
    ActiveSheet.Paste
    With Selection
        .Top = ActiveSheet.Range("K5").Top
        .Left = ActiveSheet.Range("K5").Left
    End With
    ActiveSheet.Paste
    With Selection
        .Top = ActiveSheet.Range("K6").Top
        .Left = ActiveSheet.Range("K6").Left
    End With
    
    
End Sub


```
#### どちらが良いのか？？

上のコードを見るとわかるように、  
セル選択の後にペースト実行する方法①のほうが簡潔です。

しかし、方法②のほうが適している場合もあります。  
それは、次のような場合です。  
- シートにセル選択時イベントが設定されており、図形ペースト時にイベントを発火させたくないとき
- シート保護機能により、セルの選択が禁止されているとき

## デモ

次のように、笑顔マーク、星マーク、禁止マークを  
下に続く２行にコピペしたいとします。

![マクロ実行前](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1640865306/learnerBlog/vba-question-004-copypaste-shapes/%E3%82%B9%E3%82%AF%E3%83%AA%E3%83%BC%E3%83%B3%E3%82%B7%E3%83%A7%E3%83%83%E3%83%88_2021-12-30_204003_zu8kcj.png)

上記で記載した「ペースト先のセルを選択してペースト」関数を実行すると、  
各マークが規則正しくコピペされ配置されます。

![マクロ実行後](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1640865307/learnerBlog/vba-question-004-copypaste-shapes/%E3%82%B9%E3%82%AF%E3%83%AA%E3%83%BC%E3%83%B3%E3%82%B7%E3%83%A7%E3%83%83%E3%83%88_2021-12-30_204053_nsv54q.png)

## 終わりに

VBAマクロで図形をいじったりするのは、  
マクロでしっかりしたツールを作るときよりも、  
むしろ簡単な作業の自動化のために突発的に作成する場合が多いかと思います。

そのような場合でも、図形の取り扱いを慣れておけば  
手作業よりもずっと効率よくマクロを組むことが可能となるでしょう。