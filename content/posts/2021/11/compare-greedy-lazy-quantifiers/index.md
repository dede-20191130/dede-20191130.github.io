---
title: "[JavaScript] 正規表現の実行速度検証デモ：量指定子に「Greedyな検索」と「Lazyな検索」を指定する場合"
author: dede-20191130
date: 2021-11-03T22:20:40+09:00
slug: compare-greedy-lazy-quantifiers
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
    Javascript（ブラウザ環境）で、正規表現の実行速度を量指定子のタイプの観点で調査した。<br>
    Greedyな検索とLazyな検索のいずれも利点があるため、<br>検査する文字列に依存して最適な正規表現を適用する必要がある。
{{< /box-with-title >}}

[こちらの記事](https://blog.stevenlevithan.com/archives/greedy-lazy-performance)で紹介されていた正規表現のパフォーマンス比較について、  
現在のブラウザのJavascriptではどの程度の違いが発生するのか検証してみたくなり、  
どうせならということで簡単なデモを作成した。

## 環境

ブラウザはChromeで検証した。（バージョン: 95.0.4638.54）  
*Chromium*を採用しているモダンブラウザならおおよそ同様の結果かと思われる。

## 正規表現のパフォーマンス
### ABOUT

プログラミングにおける他の実装にも言えることだが、  
正規表現において、ある目的を達成するための表現が複数ある場合がある。

例えば、次のコードはどちらも同じ出力結果となる。

```js
console.log("this is an apple".match(/\w+/g)); // [ 'this', 'is', 'an', 'apple' ]
console.log("this is an apple".match(/\b[^\s]+?\b/g)); // [ 'this', 'is', 'an', 'apple' ]
```

小さな文字列に対する置換処理や抽出処理などでは  
正規表現の表現の違いによるパフォーマンス差は出ないと思うが、  
対象文字列が肥大化するとどの程度に差が出るのだろうか？

[こちらの記事](https://blog.stevenlevithan.com/archives/greedy-lazy-performance)で論じられているケースに焦点を合わせて  
考えてみたい。

### Greedyな検索とLazyな検索

出典：[https://blog.stevenlevithan.com/archives/greedy-lazy-performance](https://blog.stevenlevithan.com/archives/greedy-lazy-performance)

量指定子（*Quantifier*）を用いた正規表現の繰り返しでは、  
貪欲な（*Greedy*）マッチングと怠惰な（*Lazy*）マッチングという概念がある。

端的に言うと、  
前者はできるだけ長い文字列でヒットさせようとして、  
後者はできるだけ短くマッチするように正規表現エンジンが文字を探索する。

■出典元で良い画像があったためそちらを参照したい。

`<0123456789>`という文字列に対して、`<.*>`というマッチング（Greedyなマッチング）をした場合、  
正規表現エンジンは次のように文字を探索する。

![Greedyなマッチング](./img01.png)

「`<`」をマッチさせたあと、  
ドット演算子（改行以外の任意の文字）とマッチする文字のかたまり（`0123456789>`）を読み込み、  
その後、`*`以降のマッチングのために来た道を戻り始め、「`>`」を見つけて終了する。

一方、`<.*>`というマッチング（Greedyなマッチング）をした場合、  
正規表現エンジンは次のように文字を探索する。


![Lazyなマッチング](./img02.png)

この場合、「`<`」をマッチさせたあと、  
ドット演算子（改行以外の任意の文字）とマッチする文字である`0`を読み込んだあと、  
すぐさま`*`以降のマッチングを検証する。  

もちろん「`>`」にはマッチしないため、再度`0`に戻り、  
次に`1`を読み込み、同じように試す。  
最終的に`9`で求めるマッチングが見つかるため、終了する。

このようにして違いが生まれる。


### パフォーマンスの比較

上記の例だと、  
Greedyなマッチングのほうが優れているように見えるが、  
もちろんそれはケースバイケースであり、  
Lazyが最適なケースも多々ある（検索エンジンで多数ヒットするかと思う）。

今回はパフォーマンスに観点を置き、  
Greedyで一気に読み込んだ場合、Lazyで一つ一つ読み込んだ場合、  
さらに、Greedyで一気に読み込んだあとに、  
`*`以降のマッチングのために戻らなければならない文字数が非常に多い場合について検証したい。

## デモの作成
### ABOUT

対象：次の形式の文字列  
`<ABC1234（山括弧以外の文字の連続）.....>ABC1234.....`

閉じる山括弧（`>`）のあとにも延々と文字が続いているため、  
Greedyな検索をする場合は、正規表現に工夫をしないと時間がかかることが予想される。

### コード

次のように、  
三種類の正規表現の実行時間を測定し、テーブルに書き出した。  

- (i) Greedyな量指定子（string.match(/<.*>/);）
- (ii) Greedyな量指定子　ドットではなく、繰り返し文字列の範囲を明示（ string.match(/<[^>]*>/);）
- (iii) Lazyな量指定子（string.match(/<.*?>/);）

```js
function compare3Type() {

    const resutlArr = [];

    let stt = (Date.now());
    for (let index = 0; index < counter; index++) {
        let x = `<${"AB12".repeat(500)}>${"C3".repeat(500)}`
        let y = x.match(/<.*?>/);

    }
    let lst = (Date.now());
    resutlArr.push(lst - stt);

    stt = (Date.now());
    for (let index = 0; index < counter; index++) {
        let x = `<${"AB12".repeat(500)}>${"C3".repeat(500)}`
        let y = x.match(/<[^>]*>/);

    }
    lst = (Date.now());
    resutlArr.push(lst - stt);

    stt = (Date.now());
    for (let index = 0; index < counter; index++) {
        let x = `<${"AB12".repeat(500)}>${"C3".repeat(500)}`
        let y = x.match(/<.*>/);

    }
    lst = (Date.now());
    resutlArr.push(lst - stt);

    return resutlArr;

}
```


### デモ

<iframe src="https://codesandbox.io/embed/compare-greedy-lazy-quantifiers-zdcml?fontsize=14&hidenavigation=1&module=%2Fcompare-worker.js&theme=dark"
     style="width:100%; height:500px; border:0; border-radius: 4px; overflow:hidden;"
     title="compare-greedy-lazy-quantifiers"
     allow="accelerometer; ambient-light-sensor; camera; encrypted-media; geolocation; gyroscope; hid; microphone; midi; payment; usb; vr; xr-spatial-tracking"
     sandbox="allow-forms allow-modals allow-popups allow-presentation allow-same-origin allow-scripts"
   ></iframe>


### 結果

ミリ秒単位で比較すると実行速度順に次の順序がある。  

**(ii) > (i) > (iii)**

これは、  
(ii)の場合、`*`以降のマッチングのために戻らなければならない文字数がゼロであるため、  
よくチューニングされていることがわかる。

一方、戻らなければならない文字数が膨大である(iii)は、  
パフォーマンスとしては最も悪い。

(i)のLazyな検索は、  
出会った文字列ごとに、一つ一つ`>`があとに続くことを検証するため遅い。

## 最適な正規表現は？

ドット演算子を用いた  
不必要に柔軟すぎるマッチング検索は避けるべき、と  
[こちらの記事](https://blog.stevenlevithan.com/archives/greedy-lazy-performance)のまとめ節でも論じられている。

よりよいパフォーマンスの正規表現をもちいるために、  
具体化すべきところを詳細に記述することが求められるようだ。
