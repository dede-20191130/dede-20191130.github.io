---
title: "[Javacript] 画像をワンクリックで別タブで開けるようにするブックマークレットを作成する"
author: dede-20191130
date: 2021-11-01T13:36:23+09:00
slug: bookmarklet-img-other-tab
draft: false
toc: true
featured: false
tags: ["JavaScript"]
categories: ["プログラミング"]
archives:
    - 2021
    - 2021-11
---

## この記事について

{{< box-with-title title="かんたんな概要と結論" >}} 
    ページ上で、リンクが貼られていない画像を検知して<br>
    別タブで開けるようにするブックマークレットを作成した。
{{< /box-with-title >}}


## ブックマークレットとは？

ユーザーがウェブブラウザのブックマークなどから起動し、  
ウェブブラウザで簡単な処理を行う簡易的なプログラム。

[こちらの記事が参考になります](https://qiita.com/aqril_1132/items/b5f9040ccb8cbc705d04)

## やりたいこと

### 事象

[（こちらの記事のページを基にご説明させていただきます。）](https://data-viz-lab.com/excel-analyticstool-intro)

ブラウザで、開いたページにいくつかの画像があることを考える。

![ページ上の画像](./bookmarklet-img-exists.png)

上の記事のように、  
画像が親のブロック要素の横幅によって圧縮されて、  
小さくて見づらい場合がある。

そのようなときに、  
別のタブとして簡単にもとのサイズの画像を開いて  
細かい部分を見たいという要求が生まれることもあるだろう。

### 方法

そのときの画像周りのDOM構造によって、画像をワンクリックで別タブで開けるかどうかが決まる。

もし、下記画像のように*anchor*タグで*img*タグが囲まれていない場合、  
ワンクリックで開くことはできない。

![別タブで開けない](./bookmarklet-img-not-open.png)

*anchor*タグで*img*タグが囲まれており、  
なおかつ*anchor*タグの*href*が画像のsrcと同じパスを与えられているならば、    
画像をワンクリックで別タブで開くことができる。

![別タブで開ける](./bookmarklet-img-open.png)

今回は、  
ブックマークレットで簡単なスクリプトを走らせて、  
すべての*anchor*タグで囲まれていない*img*タグの親要素に、  
適切な*anchor*タグを挿入していきたい。

## 環境

Chromeでの動作を確認したが、  
おそらく他のモダンブラウザで動作すると思う。

## ブックマークレットの作成
### コード

上で書いたように、  
*anchor*タグで囲まれていない*img*タグをDOM内部に持たせるための*anchor*タグを新しく作成し、  
*href*属性にsrcと同じパスをもたせる。

```js
javascript: (function () {
	const anchoredImgs = Array.from(document.querySelectorAll("a img"));
	for (const img of document.querySelectorAll("img")) {
		if (anchoredImgs.includes(img)) continue;
		const anchor = document.createElement("a");
		anchor.href = img.src;
		anchor.target = "_blank";
		img.before(anchor);
		anchor.append(img);
	}
}());
```

### 設定

それぞれのブラウザのブックマークに任意の名前で  
上のコードを登録する。

### 実行する

ブックマークからリンクを開いてスクリプトを実行。

30個ほどの*img*要素の場合、  
実行は一瞬で終了し、  
新しく*anchor*として認識された画像をクリックすると別タブで開くことができる。

![開いた画像](./bookmarklet-after-execution.png)

