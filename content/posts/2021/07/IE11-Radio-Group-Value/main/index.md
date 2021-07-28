---
title: "[Javascript] IE11でRadioボタングループのvalueを取得できない事象＆対策"
author: dede-20191130
date: 2021-07-27T22:52:15+09:00
slug: IE11-Radio-Group-Value
draft: false
toc: true
featured: false
categories: ["プログラミング"]
tags: ["Javascript","フロントエンド"]
archives:
    - 2021
    - 2021-07
---

## この記事について

Javascriptを書き、各ブラウザ対応をしたい（特にOldブラウザで動くようにしたい）場合、  
[Babel](https://babeljs.io/docs/en/)を導入・実行することで  
トランスパイルにより構文や関数を汎用的なものに変換する方法がデファクトスタンダートかと認識している。

変換の処理は開発者の手を煩わすことなく短時間で済むのが**Babel**の強力さだが、  
時折**Babel**の変換のスコープ外であるような機能があり、手動で対応策を探らなければならない。

IE11において、  
FormのRadioボタングループを取得し、そのvalueプロパティを参照した際に、  
想定した値が取得できなかったため、  
事象と対策について記したい。

## 実行環境

- ツール
    - @babel/core:7.14.6
    - @babel/preset-env:7.14.7
    - core-js:3.15.2
    - regenerator-runtime:0.13.9
- ブラウザ
    - chrome:92.0.4515.107
    - IE11:20H2(OS Build 19042.1110)


## 事象

### About

次のようなRadioボタンを含むシンプルなフォームがある。

```html
<form id="my-form" name="my-form" onsubmit="return false;">
    <h3>県庁所在地の表示</h3>
    <div>
        <input id="radio01-toyama" type="radio" name="radio01" value="富山市">
        <label for="radio01-toyama">富山県</label><br>
        <input id="radio01-ishiakwa" type="radio" name="radio01" value="金沢市">
        <label for="radio01-ishiakwa">石川県</label><br>
        <input id="radio01-fukui" type="radio" name="radio01" value="福井市">
        <label for="radio01-fukui">福井県</label>

    </div>
    <div id="result"></div>
    <div id="button-container">
        <button id="button01">表示</button>

    </div>
</form>
```

次のようなJSコードをページで実行するとする。

```js
// ...
// FormのRadioボタングループを取得し、valueを参照
// = Radioボタンのチェックされたボタンのvalue属性を参照することと同等
const radio01Value = document.forms["my-form"].radio01.value;
// ...
```

Babelによるトランスパイル後、  
ブラウザにおいて動作確認すると、  
chrome他モダンブラウザでは想定通りの挙動だが、  
{{< colored-span color="#fb9700" >}}IE11{{< /colored-span >}}ではどうもうまくいかない。

### デモページ

こちらに
[デモページ]({{< ref "../demo-ng/index.md" >}}  )
があります。

対応ブラウザでは、選択したラジオに応じた県庁所在地が表示されます。

## 何が起こっているか？

`form.radioGroupName`の呼び出しで得られる値の型の違い。  
モダンブラウザではRadioNodeListであるが、IE11ではHTMLCollection となるようだ。

```js
// モダンブラウザ
document.forms["my-form"].radio01 instanceof RadioNodeList === true
// IE11
document.forms["my-form"].radio01 instanceof HTMLCollection === true

```

ゆえに、  
上記事象ではIE11では`HTMLCollection.prototype.value`という未定義プロパティを参照するため、想定外の挙動となる。

## 対応策

いくつかあるかと思う。

### prototypeに対して挙動を定義しモダンブラウザに寄せる

ロードしたときの実行スクリプトのどこかで  
次のように定義しておく。


```js
Object.defineProperty(HTMLCollection.prototype, "value", {
    get: function () {
        for (let idx = 0; idx < this.length; idx++) {
            if (this[idx].checked) return this[idx].value;
        }
        return undefined;
    },
});

```

#### Pros（長所）

定義しておけば、  
以降のRadioボタンまわりのコードを書き換える必要がない。

#### Cons（短所）

コメントで実装の意味を記しておかないと  
実装者以外が意図を測りかねるかもしれない。

### CSSセレクターで要素を取得する

jQueryを用いるか、querySelectorAPIを用いて  
CSSのセレクター構文で要素を指定して、チェックされたボタンのvalue属性を直接取得するようにする。

[jQuery]
```js
const radio01Value = $("form[name='my-form'] input[type='radio']:checked").val();
```

[プレーンJS]
```js
const radio01Value = document.querySelector("form[name='my-form'] input[type='radio']:checked").value;
```

#### Pros

シンプル。  
変換前の書き方と同じように一行で書き下すことができる。

#### Cons

書き換えコストは少ないものの、  
Radioボタンまわりのコードを書き換える必要がある。

## まとめ

原因はコレクション参照時のオブジェクトの型の違い。  
対応策はいくつかあるが、querySelectorで一行で置き換えるのが良いかと思う。



