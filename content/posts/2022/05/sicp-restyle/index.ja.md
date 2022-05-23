---
title: "SICP教本をよりReadableにするためのScriptAutoRunnerスクリプト"
author: dede-20191130
date: 2022-05-22T20:03:42+09:00
slug: sicp-restyle
draft: false
toc: true
featured: false
tags: ["javascript","Chrome"]
categories: ["プログラミング"]
archives:
    - 2022
    - 2022-05
---

## この記事について

{{< box-with-title title="かんたんな概要と結論" >}} 
    CDNからライブラリを読み込むことでハイライトされていないコードに対して<br/>
    シンタックスハイライティングを行うことが可能。<br/>
    コードの実行環境にはScriptAutoRunnerを使用。
{{< /box-with-title >}}

こんにちは、dedeです。

この記事では、  
Webで一般公開されているSICP教本について  
スクリプトを用いてより見やすくする方法について紹介します。

## SICP教本とは
### ABOUT

SICPは、コンピュータ科学における古典的な名著の一つです。  
再帰的手続きやモジュール化の利益など、コンピュータプログラミングの頻出パターンについて説明する本です。

### 公開サイト

[こちら](https://mitpress.mit.edu/sites/default/files/sicp/full-text/book/book.html)で書籍版と同等の内容のテキストが公開されています。

### 問題点

上記サイトですが、個人的見解として二点見にくい点がありました。

- 各内容の説明のために、 _Scheme_ （プログラミング言語Lispのバリエーションの一つ）を用いているが、  
それらのコードブロックにはシンタックスハイライトが施されていないため、少々見にくい。
- 脚注部分の文字サイズがやや小さい。

これらを改善するために、ブラウザで動作する、デザインを改善するスクリプトを書きました。  

通常、ブラウザで自作Javascriptコードを動作させるには  
「開発者コンソールで対話的に実行する」あるいは「ブックマークレットとして登録し実行する」のですが、  
両者ともページを更新するたびに再度実行しなければならない欠点がありました。

しかし、Chromeブラウザの場合、便利な拡張機能があります。

## ScriptAutoRunnerとは

[[ScriptAutoRunner](https://chrome.google.com/webstore/detail/scriptautorunner/gpgjofmpmjjopcogjgdldidobhmjmdbm?hl=ja)

あらかじめ特定ドメインと流したいスクリプトコードのペアを登録し、  
該当するページを開く/更新するたびにコードが実行されるようにできる便利なChrome拡張です。

![ScriptAutoRunner（公式より）](https://lh3.googleusercontent.com/LUHrciH1gr-dNe_0yrVuje-TYIb66LIJePum2HDipQ8HFPB_kjpvQqLnYxbw7Wn_drDTLf7l604zciVYugAUvg6ic00=w640-h400-e365-rj-sc0x00ffffff)

## 作成環境・利用ツール

- Google Chrome バージョン: 101.0.4951.67
- highlight.js v11.5.1

## スクリプトのコード
### 全体

全体のコードです。
続くセクションで個々の意味と動作について記します。

```js
// if document is not fully loaded, load event take the place
if (document.readyState === "complete") {
    restyleSICP();
} else {
    window.addEventListener("load", restyleSICP);
}

function restyleSICP() {
    // only in specific paths this script runs
    if (!window.location.pathname.startsWith("/sites/default/files/sicp/full-text/book")) return;
    highlightSchemeCode();
    expandFootnote();
}

function highlightSchemeCode() {
    // exclude elements which contain some img elements
    const targetTts = Array.from(document.querySelectorAll("p > tt:first-child:last-child")).filter(e => !e.querySelector("img"));

    for (const tt of targetTts) {
        const pre = document.createElement("pre");
        const code = document.createElement("code");

        code.classList.add("language-scheme");
        code.innerHTML = tt.innerHTML;

        // replace tt element with pre + code elements
        pre.append(code);
        tt.before(pre);
        tt.remove();
    }

    const link = document.createElement("link");
    link.rel = "stylesheet";
    link.href = "https://cdnjs.cloudflare.com/ajax/libs/highlight.js/11.5.1/styles/a11y-dark.min.css";

    const hlScr = document.createElement("script");
    hlScr.src = "//cdnjs.cloudflare.com/ajax/libs/highlight.js/11.5.1/highlight.min.js";
    hlScr.onload = () => {
        hljs.configure({
            ignoreUnescapedHTML: true
        });
        const schemeScr = document.createElement("script");
        schemeScr.src = "https://cdnjs.cloudflare.com/ajax/libs/highlight.js/11.5.1/languages/scheme.min.js";
        schemeScr.onload = () => hljs.highlightAll();
        document.head.append(schemeScr);

    }
    document.head.append(link, hlScr);


}

function expandFootnote() {
    document.querySelector(".footnote").style.fontSize = "1rem";
}
```

### 始動

```js
// if document is not fully loaded, load event take the place
if (document.readyState === "complete") {
    restyleSICP();
} else {
    window.addEventListener("load", restyleSICP);
}
```

DOM読み込みが完了していない場合と既に完了している場合とで  
メイン処理の実行タイミングが異なります。

### 実行ページのチェック

```js
function restyleSICP() {
    // only in specific paths this script runs
    if (!window.location.pathname.startsWith("/sites/default/files/sicp/full-text/book")) return;
    // ...
}
```

ScriptAutoRunnerではドメインの指定は可能でもパスの指定はできないため、  
SICPテキストに関係したページでのみ実行継続するようにします。

### ハイライティング

```js
function highlightSchemeCode() {
    // exclude elements which contain some img elements
    const targetTts = Array.from(document.querySelectorAll("p > tt:first-child:last-child")).filter(e => !e.querySelector("img"));

    for (const tt of targetTts) {
        const pre = document.createElement("pre");
        const code = document.createElement("code");

        code.classList.add("language-scheme");
        code.innerHTML = tt.innerHTML;

        // replace tt element with pre + code elements
        pre.append(code);
        tt.before(pre);
        tt.remove();
    }

    
```

各ページで _Scheme_ のコードに相当する部分はtt要素なので、  
それらを探し出し、pre要素＋code要素に置き換えます。  
pre要素＋code要素は、 _highlight.js_ の検知対象です。

img要素が含まれるコードブロックもあり、ハイライトするとかえって見づらくなるので除外します。

```js
const link = document.createElement("link");
    link.rel = "stylesheet";
    link.href = "https://cdnjs.cloudflare.com/ajax/libs/highlight.js/11.5.1/styles/a11y-dark.min.css";

    const hlScr = document.createElement("script");
    hlScr.src = "//cdnjs.cloudflare.com/ajax/libs/highlight.js/11.5.1/highlight.min.js";
    hlScr.onload = () => {
        hljs.configure({
            ignoreUnescapedHTML: true
        });
        const schemeScr = document.createElement("script");
        schemeScr.src = "https://cdnjs.cloudflare.com/ajax/libs/highlight.js/11.5.1/languages/scheme.min.js";
        schemeScr.onload = () => hljs.highlightAll();
        document.head.append(schemeScr);

    }
    document.head.append(link, hlScr);


}
```

シンタックスハイライト実施には、 _highlight.js_ を利用します。

CDNから _highlight.js_ のメインスクリプトおよび _Scheme_ のための付属スクリプトを順番に読み込み、  
すべて読み込み完了したあとに`highlightAll`を実行します。


### 脚注スタイル

```js
function expandFootnote() {
    document.querySelector(".footnote").style.fontSize = "1rem";
}
```

脚注の文字サイズを大きくします。
## デモ

![before](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1653230225/learnerBlog/sicp-restyle/mitpress.mit.edu_sites_default_files_sicp_full-text_book_book-Z-H-12.html_aiwgpd.png)

![after](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1653230225/learnerBlog/sicp-restyle/mitpress.mit.edu_sites_default_files_sicp_full-text_book_book-Z-H-12.html_1_demakx.png)

## 終わりに

ScriptAutoRunnerは最近知りました。  
ブックマークレットを実行するよりも手軽かつメンテナンス容易にJavascriptコードを実行できるので、  
サイトを見やすくする際に、様々に役に立つ拡張かと思います。
