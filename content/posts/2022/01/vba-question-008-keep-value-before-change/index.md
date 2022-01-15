---
title: "[教えて！VBA] 第8回 変更前のセルの値を保持/利用するにはどうすればいいの？？"
author: dede-20191130
date: 2022-01-06T14:10:19+09:00
slug: vba-question-008-keep-value-before-change
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


![](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1641629426/learnerBlog/vba-question-008-keep-value-before-change/keep-value-before-change_bdf1ed.png)

## この記事について

{{< box-with-title title="かんたんな概要と結論" >}} 
    VBAマクロで、特定の操作で処理を実行する仕組みをイベントと呼びます。<br/>
    イベントを利用し、セル入力時に処理を実行させて、変更前の値を取り扱うことができます。
{{< /box-with-title >}}

こんにちは、dedeです。

この記事では、  
VBAマクロに関する質問のうち、  
皆が疑問に思っているトピックについて解説いたします。

今回は、
{{< colored-span color="#fb9700" >}}変更前のセルの値を記録・保持して利用したり、変更値をチェックして入力前の値に戻したりする方法{{< /colored-span >}}  
を解説いたします。

※この記事は、Office VBAマクロのうち  
Excel VBAマクロに関するトピックです。

レベル：<b>中級者向け</b>

## やりたいこと

シートのセルにデータを入力すると、  
通常は変更前の古いデータは消去されます（「元に戻す」ボタンで戻すことはできますが）。

入力したタイミングに応じて、変更前の値をシートのどこかに移したい場合は、  
VBAマクロの出番です。

![変更前の値をシートのどこかに移す](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1641629426/learnerBlog/vba-question-008-keep-value-before-change/keep-value-before-change2_pmfntt.png)

このとき、  
{{< colored-span color="#fb9700" >}}イベント{{< /colored-span >}}と呼ばれるVBAマクロの仕組みを利用することになります。

また、変更したデータを入力した時に、  
適切ではない値だった場合に自動でチェックして前の値に戻したいときもあるでしょう。

![入力値の自動チェック](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1641629426/learnerBlog/vba-question-008-keep-value-before-change/keep-value-before-change3_ev7tzv.png)

そのような各パターンについて見ていきます。
## パターン① 変更前の値を他のセルに移したい
### ABOUT

入力するたびに、  
前の値をシートのどこかのセルに記録していくようなマクロを作成したいという希望があるとしましょう。

「入力するたび」というタイミングの検知のために、  
イベントという仕組みを利用します。

### イベントとは？

ブックやシートに対して、  
ある特定の操作が行われた際に自動でVBAマクロが発火（処理の開始）されるような仕組みです。

例としては次のようなものがあります

<table>
  <thead>
    <tr>
      <th>対象</th>
      <th>イベント名</th>
      <th>内容</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td rowspan="2">ブック<br/>（WorkBook）</td>
      <td>Open</td>
      <td>ファイルを起動して<br />ブックが開かれたときに実行</td>
    </tr>
    <tr>
      <td>BeforeClose</td>
      <td>閉じるボタンなどでブックが閉じられる時に実行</td>
    </tr>
    <tr>
      <td rowspan="2">シート<br/>（Worksheet）</td>
      <td>SelectionChange</td>
      <td>現在のセルから別のセルを選択した時に実行</td>
    </tr>
    <tr>
      <td>Change</td>
      <td>セルの内容を変更した時に実行</td>
    </tr>
  </tbody>
</table>


### コード

イベントの実装は、  
イベントプロシージャというSubプロシージャの記入により行います。

イベントプロシージャは、  
イベントを設定したいシートのコードウインドウ上に記入します。

![マクロ記述シートの指定](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1641629426/learnerBlog/vba-question-008-keep-value-before-change/%E3%82%B9%E3%82%AF%E3%83%AA%E3%83%BC%E3%83%B3%E3%82%B7%E3%83%A7%E3%83%83%E3%83%88_2022-01-08_123712_sde8dn.png)

![シートのコードウインドウ](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1641629426/learnerBlog/vba-question-008-keep-value-before-change/%E3%82%B9%E3%82%AF%E3%83%AA%E3%83%BC%E3%83%B3%E3%82%B7%E3%83%A7%E3%83%83%E3%83%88_2022-01-08_123919_ay4oqw.png)

変更前の値を他のセルに移すマクロのコードは下記です。

```vb
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim OldValue As Variant
    Dim currentCellAddress
    
    '//ポイント①　変更前の値を調べたいセル以外では処理を行わないようにする
    If Target.Address(False, False) <> "B3" Then Exit Sub
    
    '//ポイント②　いったんイベントを無効化し、イベントの発火を停止する
    Application.EnableEvents = False
    
    '//カーソル位置を退避させる
    currentCellAddress = Selection.Address
    
    '//ポイント③　いったんセルの値をもとに戻し、古い値を変数に退避させてから再度変更後の値に戻す
    Application.Undo
    OldValue = Target.Value
    Application.Undo
    
    '//変更前の値を別のセルに移す
    Me.Range("D3").Value = OldValue
    
    '//カーソル位置を元に戻す
    Me.Range(currentCellAddress).Select
    
    '//イベントを再度有効にする
    Application.EnableEvents = True
    
End Sub

```

ポイントは三点あり、  
①変更前の値を調べたいセル以外では処理を行わないようにするため、`Target`（変更したセルのRangeを参照する引数）のアドレスについて絞り込みを行います。  
上では、B3セル以外では処理を行わないようにします。  

②処理の最中に他のイベントが発火しないように、`EnableEvents`を変更しておきます。  

③いったんセルの値をもとに戻し、古い値を変数に退避させてから再度変更後の値に戻します。

### デモ

変更前に東京都が入っています。

![セル入力前](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1641629426/learnerBlog/vba-question-008-keep-value-before-change/%E3%82%B9%E3%82%AF%E3%83%AA%E3%83%BC%E3%83%B3%E3%82%B7%E3%83%A7%E3%83%83%E3%83%88_2022-01-08_131501_xwy5gu.png)

新しいデータを入力後、別のセルに東京都の文字が移されます。

![セル入力後](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1641629427/learnerBlog/vba-question-008-keep-value-before-change/%E3%82%B9%E3%82%AF%E3%83%AA%E3%83%BC%E3%83%B3%E3%82%B7%E3%83%A7%E3%83%83%E3%83%88_2022-01-08_131551_rsnh6r.png)

## パターン② 変更前の値を他のセルに移し、削除した値を別のセルに移したい
### ABOUT

変更ではなく、  
削除した場合（空欄として入力した場合）には別のセルに値を移すという仕様にしたいとしましょう。

その場合、いくらか条件分岐が増え、  
また、削除したことを示すフラグを変数に持たせる必要が発生します。

### コード

```vb
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim inputValue As Variant
    Dim OldValues(0 To 1) As Variant
    Dim currentCellAddress
    
    '//変更前の値を調べたいセル以外では処理を行わないようにする
    If Target.Address(False, False) <> "B3" Then Exit Sub
    
    '//いったんイベントを無効化し、イベントの発火を停止する
    Application.EnableEvents = False
    
    '//カーソル位置を退避させる
    currentCellAddress = Selection.Address
    
    '//いったんセルの値をもとに戻し、古い値を変数に退避させてから再度変更後の値に戻す
    inputValue = Target.Value                    '//入力した値を変数に退避
    Application.Undo
    '//ポイント①　入力した値が空欄であるかどうかで削除処理か変更処理かを分ける
    If inputValue = "" Then
        OldValues(0) = "delete"
    Else
        OldValues(0) = "change"
    End If
    OldValues(1) = Target.Value
    Application.Undo
    
    '//変更前の値を別のセルに移す
    '//ポイント②　削除処理であればF3セルに、変更処理であればD3セルに移す
    If OldValues(0) = "delete" Then
        Me.Range("F3").Value = OldValues(1)
    Else
        If OldValues(1) <> "" Then Me.Range("D3").Value = OldValues(1)
        
    End If
    
    '//カーソル位置を元に戻す
    Me.Range(currentCellAddress).Select
    
    '//イベントを再度有効にする
    Application.EnableEvents = True
    
End Sub
```

ポイントは、  
①入力した値が空欄であるかどうかで、  
削除であるか変更であるか判別できるようなフラグを持っておきます。

②削除処理、変更処理ごとに反映するセルを分けています。

### デモ

県名を変更すると、  
「前回の入力値」欄に変更前の値が入ります。

![セル入力前](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1641629427/learnerBlog/vba-question-008-keep-value-before-change/%E3%82%B9%E3%82%AF%E3%83%AA%E3%83%BC%E3%83%B3%E3%82%B7%E3%83%A7%E3%83%83%E3%83%88_2022-01-08_135817_ckznao.png)
![セル入力後](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1641629427/learnerBlog/vba-question-008-keep-value-before-change/%E3%82%B9%E3%82%AF%E3%83%AA%E3%83%BC%E3%83%B3%E3%82%B7%E3%83%A7%E3%83%83%E3%83%88_2022-01-08_135842_jz3epi.png)

また、データ削除すると、  
「削除された値」欄に削除前のデータが入ります。  
前回の入力値欄のデータには変更がありません。

![セルの値の削除後](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1641629427/learnerBlog/vba-question-008-keep-value-before-change/%E3%82%B9%E3%82%AF%E3%83%AA%E3%83%BC%E3%83%B3%E3%82%B7%E3%83%A7%E3%83%83%E3%83%88_2022-01-08_135854_nyajbo.png)

## パターン③ 変更時に値をチェックし、不正な値であれば入力前に戻したい
### ABOUT

イベントの仕組みを使って、  
入力した際に即座に入力値のチェックをして不正な値を弾きたい時の実装を記します。

### コード

```vb

Private Const HOKURIKU As String = "富山;石川;福井;"

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim OldValue As Variant
    Dim currentCellAddress
    Dim HOKURIKUs() As String
    Dim oneHokuriku As Variant
    Dim isHokuriku As Boolean
    
    '//変更前の値を調べたいセル以外では処理を行わないようにする
    If Target.Address(False, False) <> "B3" Then Exit Sub
    
    '//いったんイベントを無効化し、イベントの発火を停止する
    Application.EnableEvents = False
    
    '//カーソル位置を退避させる
    currentCellAddress = Selection.Address
    
    '//北陸県を配列として取得する
    HOKURIKUs = Split(HOKURIKU, ";")
    isHokuriku = False
    
    '//ポイント①　入力値が北陸県のいずれかに該当すれば戻さない
    '//　　　　　　該当しなければ変更前に戻すようにする
    For Each oneHokuriku In HOKURIKUs
        If Target.Value = oneHokuriku Then
            isHokuriku = True
            Exit For
        End If
    Next oneHokuriku
    If Not isHokuriku Then Application.Undo
    
    '//カーソル位置を元に戻す
    Me.Range(currentCellAddress).Select
    
    '//イベントを再度有効にする
    Application.EnableEvents = True
    
End Sub
```

ポイントとしては、  
入力値チェックの機構を導入しています。

入力値が北陸県（富山;石川;福井）のいずれかに該当すればそのまま入力を続行しますが、  
もし北陸県以外の文字列が入力されれば、変更前に戻します。

### デモ

入力したときに北陸県以外の文字であれば、  
空欄に戻されます。

![セル入力前](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1641629425/learnerBlog/vba-question-008-keep-value-before-change/%E3%82%B9%E3%82%AF%E3%83%AA%E3%83%BC%E3%83%B3%E3%82%B7%E3%83%A7%E3%83%83%E3%83%88_2022-01-08_143057_kjrtk5.png)
![東京都と入力中](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1641629425/learnerBlog/vba-question-008-keep-value-before-change/%E3%82%B9%E3%82%AF%E3%83%AA%E3%83%BC%E3%83%B3%E3%82%B7%E3%83%A7%E3%83%83%E3%83%88_2022-01-08_143129_lrviqh.png)
![入力後空欄に戻る](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1641629425/learnerBlog/vba-question-008-keep-value-before-change/%E3%82%B9%E3%82%AF%E3%83%AA%E3%83%BC%E3%83%B3%E3%82%B7%E3%83%A7%E3%83%83%E3%83%88_2022-01-08_143145_eqbtq5.png)

北陸県のどれかであれば、  
チェックが通り、通常通り入力が完了できます。

![北陸県ならば変更可能](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1641629425/learnerBlog/vba-question-008-keep-value-before-change/%E3%82%B9%E3%82%AF%E3%83%AA%E3%83%BC%E3%83%B3%E3%82%B7%E3%83%A7%E3%83%83%E3%83%88_2022-01-08_143218_ry0idh.png)

## パターン④ 変更した値を複数保持し、リストとして他のセルに出力したい
### ABOUT

変更前の値として複数持てるようにしたいという希望があるとします。

次のように、一つのセルに改行で分けられたリストの形で  
最大10個の値を格納していきます。

![リストとして格納](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1641629425/learnerBlog/vba-question-008-keep-value-before-change/%E3%82%B9%E3%82%AF%E3%83%AA%E3%83%BC%E3%83%B3%E3%82%B7%E3%83%A7%E3%83%83%E3%83%88_2022-01-08_165407_ieanvr.png)

### コード

```vb
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim OldValue As Variant
    Dim currentCellAddress
    Dim lists() As String
    Dim firstCount As Long
    Dim valueFromList As Variant
    Dim i As Long
    
    '//変更前の値を調べたいセル以外では処理を行わないようにする
    If Target.Address(False, False) <> "B3" Then Exit Sub
    
    '//いったんイベントを無効化し、イベントの発火を停止する
    Application.EnableEvents = False
    
    '//カーソル位置を退避させる
    currentCellAddress = Selection.Address
    
    '//いったんセルの値をもとに戻し、古い値を変数に退避させてから再度変更後の値に戻す
    Application.Undo
    OldValue = Target.Value
    Application.Undo
    
    '//変更前の値を別のセルに移す
    '////配列の形で現在のリストを取得する
    lists = Split(Me.Range("D3").Value, vbLf)
    firstCount = LBound(lists)
    
    '////リスト項目が10以上であれば、最初の値を消すようにする
    If (UBound(lists) - LBound(lists) + 1) = 10 Then firstCount = firstCount + 1
    
    '////新しいリストを作成する
    For i = firstCount To UBound(lists)
        valueFromList = valueFromList & lists(i) & vbLf
    Next i
    valueFromList = valueFromList & OldValue
    
    Me.Range("D3").Value = valueFromList
    
    '//カーソル位置を元に戻す
    Me.Range(currentCellAddress).Select
    
    '//イベントを再度有効にする
    Application.EnableEvents = True
    
End Sub


```

改行コードでリストを分割して配列として取り扱い、  
再度格納する際には分割したそれぞれを改行コードで結合します。

### デモ

これまでにデータ1~10を入力済みで、  
現在の値がデータ11である状態です。

![データ12の入力前](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1641629425/learnerBlog/vba-question-008-keep-value-before-change/%E3%82%B9%E3%82%AF%E3%83%AA%E3%83%BC%E3%83%B3%E3%82%B7%E3%83%A7%E3%83%83%E3%83%88_2022-01-08_165805_omxtpp.png)

あたらしく「データ12」を入力すると、  
最も古い値が消去され、  
データ2～データ11がリストに格納されています。

![データ12の入力後](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1641629426/learnerBlog/vba-question-008-keep-value-before-change/%E3%82%B9%E3%82%AF%E3%83%AA%E3%83%BC%E3%83%B3%E3%82%B7%E3%83%A7%E3%83%83%E3%83%88_2022-01-08_170133_buutln.png)

## 終わりに

イベントを使いこなすことで、  
Excel単独でも高度なアプリケーションのように  
さまざまな機能をもたせることができます。

イベントの中でも、  
`Change`イベントはよく使われるので、  
これを理解するとマクロが上達するでしょう。
