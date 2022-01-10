---
title: "[教えて！VBA] 第9回 フォントやボタンをカスタマイズできるメッセージボックスを作成するにはどうすればいいの？？"
author: dede-20191130
date: 2022-01-10T12:34:09+09:00
slug: vba-question-009-create-own-message-frm
draft: false
toc: true
featured: false
tags: ["VBA","Excel"]
categories: ["プログラミング"]
archives:
    - 2022
    - 2022-01
---


![](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1641804072/learnerBlog/vba-question-009-create-own-message-frm/create-own-message-frm_gauuok.png)

## この記事について

{{< box-with-title title="かんたんな概要と結論" >}} 
    独自のメッセージボックスを使ってカスタマイズを行うためには、<br/>
    ユーザーフォームという機能を利用する必要があります。<br/><br/>
    ユーザーフォームを使って、URLリンクを貼ったり、<br/>
    ボタンの数を増やしたりすることができます。
{{< /box-with-title >}}

こんにちは、dedeです。

この記事では、  
VBAマクロに関する質問のうち、  
皆が疑問に思っているトピックについて解説いたします。

今回は、
{{< colored-span color="#fb9700" >}}フォントやボタンをカスタマイズできるような、独自のメッセージボックスを作成する方法{{< /colored-span >}}  
を解説いたします。

※この記事は、Office VBAマクロのうち  
Excel VBAマクロに関するトピックです。

レベル：<b>中級者向け</b>

## 標準のメッセージボックスの機能

標準のメッセージボックスを利用するには、  
`VBA.Interaction` のメンバーである`MsgBox`関数を使用します。

```vb
MsgBox (prompt, [ buttons, ] [ title, ] [ helpfile, context ])
```

標準のメッセージボックスには各種の機能がありますが、  
その種類は限られています。

詳しくは、[公式](https://docs.microsoft.com/ja-jp/office/vba/language/reference/user-interface-help/msgbox-function)に記載があります。

例えば、  
次のように「はい」「いいえ」の選択肢ボタンを表示したり、  
警告アイコンやエラーアイコンを付記することはできます。

![標準のメッセージボックス](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1641804072/learnerBlog/vba-question-009-create-own-message-frm/%E3%82%B9%E3%82%AF%E3%83%AA%E3%83%BC%E3%83%B3%E3%82%B7%E3%83%A7%E3%83%83%E3%83%88_2022-01-10_133115_dhbpvp.png)

## やりたいこと

上記の標準機能では、  
次のようなことができません。

- フォントをカスタマイズする（色の変更など）
- URLリンクを貼り付ける
- 選択肢ボタンの他に、別の処理を行うボタンを配置する

これらを実現するために、  
Excel VBAのユーザーフォーム機能を利用します。

【ユーザーフォームとは？】

ボタンやテキストボックスなどのコントロールを自由に配置してカスタマイズできる画面です。  
詳細な利用方法は [こちら](https://www.239-programing.com/excel-vba/ufm/ufm012.html)

## 実装する
### レベル1. 独自のメッセージフォームを構築

まず、独自メッセージのためのフォームを作成します。

![独自メッセージのためのフォーム](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1641804072/learnerBlog/vba-question-009-create-own-message-frm/%E3%82%B9%E3%82%AF%E3%83%AA%E3%83%BC%E3%83%B3%E3%82%B7%E3%83%A7%E3%83%83%E3%83%88_2022-01-10_140838_piu7so.png)

その後、フォームにラベルやボタンなどの部品（コントロールと呼びます）を設置します。

![メッセージボックスの部品設置](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1641804072/learnerBlog/vba-question-009-create-own-message-frm/create-own-message-frm2_gcyhvh.png)

最後にフォームのコードウィンドウ上にコードを記入します。

```vb
'//OKボタン
Private Sub CommandButton_OK_Click()
    '//結果を現在表示中のシートに記載
    ActiveSheet.Range("B5").Value = "結果です。"
    ActiveSheet.Range("B6").Value = "ああいいううええおお"
    '//フォームを閉じる
    Unload Me
End Sub


'//キャンセルボタン
Private Sub CommandButton_Cancel_Click()
    '//フォームを閉じる
    Unload Me
End Sub


```

OKボタン押下時には結果をシートに記入しまてからフォームを閉じますが、  
キャンセルボタン押下時には結果処理をキャンセルし、フォームを閉じます。

以下、デモです。

1. シートに独自メッセージ呼び出し用のボタンを設置します。

![呼び出し用のボタン](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1641804072/learnerBlog/vba-question-009-create-own-message-frm/%E3%82%B9%E3%82%AF%E3%83%AA%E3%83%BC%E3%83%B3%E3%82%B7%E3%83%A7%E3%83%83%E3%83%88_2022-01-10_165010_d1mmro.png)

このボタンを押すと独自メッセージが呼び出されます。  
ボタンには次のように呼び出し用のプログラムを設定します。

```vb
Private Sub CommandButton_ShowMessage_Click()
    '//独自メッセージを表示
    UserForm_MyMessage.Show
End Sub

```

今回はデモなのでシート上のボタンの押下で呼び出しというようにしていますが、  
呼び出したい場所で`UserForm_MyMessage.Show`を記述すればよいだけなので、  
実際上はどのようなプログラムでもメッセージを表示させることができます。

2. ボタンを押下して独自メッセージを表示させます。

![独自メッセージを表示](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1641804072/learnerBlog/vba-question-009-create-own-message-frm/%E3%82%B9%E3%82%AF%E3%83%AA%E3%83%BC%E3%83%B3%E3%82%B7%E3%83%A7%E3%83%83%E3%83%88_2022-01-10_165522_qjhjlt.png)

3. メッセージボックス上のOKボタンを押下すると、シートに結果の文字を表示します。

![OKボタンを押下](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1641804073/learnerBlog/vba-question-009-create-own-message-frm/%E3%82%B9%E3%82%AF%E3%83%AA%E3%83%BC%E3%83%B3%E3%82%B7%E3%83%A7%E3%83%83%E3%83%88_2022-01-10_165533_mxajkw.png)

### レベル2. フォントをカスタマイズ

独自メッセージボックス上の文字はラベルというコントロールに記述されており、  
ラベルの文字は色を変えたり大きさを変えたりすることができます。

![フォントをカスタマイズ](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1641804073/learnerBlog/vba-question-009-create-own-message-frm/%E3%82%B9%E3%82%AF%E3%83%AA%E3%83%BC%E3%83%B3%E3%82%B7%E3%83%A7%E3%83%83%E3%83%88_2022-01-10_170315_vb4wiq.png)

「メールの添付」の文字のためのラベルを新しく挿入し、  
文字色を大きさ、太さを変更しました。

### レベル3. URLリンクを配置

ラベルには、  
URLリンクも配置することができます。

![URLリンクを配置](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1641804071/learnerBlog/vba-question-009-create-own-message-frm/%E3%82%B9%E3%82%AF%E3%83%AA%E3%83%BC%E3%83%B3%E3%82%B7%E3%83%A7%E3%83%83%E3%83%88_2022-01-10_171053_p71qee.png)

上では、URLとしてexample.comを設定します。

ラベルをクリックした際のイベントとして次のようにコードを書きます。

```vb

'//URL記載ラベルのクリックイベント
Private Sub Label_URLLink_Click()
    ThisWorkbook.FollowHyperlink Address:="http://example.com"
End Sub

```

以下、デモです。

1. メッセージボックスを表示し、URLリンクをクリックすると……

![URLリンクをクリック](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1641804072/learnerBlog/vba-question-009-create-own-message-frm/%E3%82%B9%E3%82%AF%E3%83%AA%E3%83%BC%E3%83%B3%E3%82%B7%E3%83%A7%E3%83%83%E3%83%88_2022-01-10_171121_xant0f.png)

2. example.comのサイトがブラウザで開かれます。

![リンクを開く](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1641804071/learnerBlog/vba-question-009-create-own-message-frm/%E3%82%B9%E3%82%AF%E3%83%AA%E3%83%BC%E3%83%B3%E3%82%B7%E3%83%A7%E3%83%83%E3%83%88_2022-01-10_171133_wzwsek.png)

### レベル4. ボタンを増やし、メッセージの機能を追加する

「OK」「キャンセル」「はい」「いいえ」などの選択肢ボタンの他に、  
具体的な動作を表すボタンを追加することもできます。

このようにすれば、メッセージボックスに特有の機能を持たせることができます。

![ボタン追加](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1641804071/learnerBlog/vba-question-009-create-own-message-frm/%E3%82%B9%E3%82%AF%E3%83%AA%E3%83%BC%E3%83%B3%E3%82%B7%E3%83%A7%E3%83%83%E3%83%88_2022-01-10_172723_gsvp75.png)

上の図では、現在表示中のシートではなく、  
「結果シート」というシートに文字を書き込むボタンを追加しています。

ボタン押下時のイベント処理は次のようなコードとして記述します。

```vb
'//結果シートに書きこむボタン
Private Sub CommandButton_WriteResultSheet_Click()
    '//結果を結果シートに記載
    Worksheets("結果シート").Range("B2").Value = "結果です。"
    Worksheets("結果シート").Range("B3").Value = "ああいいううええおお"
    '//フォームを閉じる
    Unload Me
End Sub

```

ボタンの数は常識的な範囲ではいくつでも配置できる（上限があったかどうかは把握しておりませんが）ので、  
実装の方法次第で、非常に多機能なメッセージボックスに仕上げることも可能です。

以下、デモです。

1. 結果シートには何も記入されていない初期状態です。

![結果シート初期状態](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1641804071/learnerBlog/vba-question-009-create-own-message-frm/%E3%82%B9%E3%82%AF%E3%83%AA%E3%83%BC%E3%83%B3%E3%82%B7%E3%83%A7%E3%83%83%E3%83%88_2022-01-10_172756_muliwb.png)

2. 「結果シートに書き込む」ボタンを押下します。

![結果シートに書き込むボタンを押下](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1641804072/learnerBlog/vba-question-009-create-own-message-frm/%E3%82%B9%E3%82%AF%E3%83%AA%E3%83%BC%E3%83%B3%E3%82%B7%E3%83%A7%E3%83%83%E3%83%88_2022-01-10_172818_oqpwmz.png)

3. OKボタンを押したときとは異なり、結果シートに結果の文字が書き込まれ、フォームが閉じます。

![結果シートに書き込み](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1641804072/learnerBlog/vba-question-009-create-own-message-frm/%E3%82%B9%E3%82%AF%E3%83%AA%E3%83%BC%E3%83%B3%E3%82%B7%E3%83%A7%E3%83%83%E3%83%88_2022-01-10_172831_xiiu1b.png)

## ソースファイルについて

デモで使用したソースファイルについて、  
[こちら](https://github.com/dede-20191130/My_VBA_Tools/tree/master/Public/2022/01/create-own-message-frm)からダウンロードできます。

## 終わりに

メッセージボックスのカスタマイズを行うと、  
注意させたい文言を赤文字で書いたり、  
業務上重要なサイトをURLリンクで貼り付けたりできるので、  
VBAマクロの便利さをより向上させることができるでしょう。