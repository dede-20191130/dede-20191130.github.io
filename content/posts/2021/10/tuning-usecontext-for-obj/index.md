---
title: "[React Hook] レンダリングのチューニング ~オブジェクトをContextとして利用する場合 & サンプルコード~"
author: dede-20191130
date: 2021-10-30T18:56:19+09:00
slug: tuning-usecontext-for-obj
draft: false
toc: true
featured: false
tags: ["JavaScript","React","フロントエンド"]
categories: ["プログラミング"]
archives:
    - 2021
    - 2021-10
---

## この記事について

{{< box-with-title title="かんたんな概要と結論" >}} 
    オブジェクトを<i>Context</i>として使用する際のレンダリングのチューニングを実行した。<br>
    外部ライブラリの<i>isDeepEqual</i>、<i>useDeepCompareEffect</i>を使用することで<br>
    良い感じにDeep Compareを実装できる。
{{< /box-with-title >}}

*React Hook*によるブラウザアプリ画面の作成の際、  
レンダリングされる*Component*の数を絞る（チューニングする）ことで  
実行されるコードのステップ数を減らし、パフォーマンスを改善することができる。

オブジェクトを*Context*として各*Component*にわたす処理のコードにおいて、  
いくつかの観点ごとにチューニングする必要があったため、  
ここにそれを記したい。


## レンダリングのチューニング

### やりたいこと

*React Hook*においてオブジェクトを*Context*として各*Component*にわたす際に、  
オブジェクトを更新するたびに、通常は*Context*を使用するすべての*Component*が再レンダリングされる。  

オブジェクトが再度同じ値に設定される場合に、  
不要なレンダリングを防ぐような仕組みが、Reactの標準APIにないかと探したが、  
どうやら存在しないようだった。

また、もう一つの課題として、  
利用先の*Component*で*Context*のオブジェクトを取得し、  
その値（参照ではなく、内容）に変更があるたびに*useEffect*を実行させたいというものがある。  
*useEffect*はimmutableな値のみを依存関係に含めることができるので、  
今回のようなオブジェクトの変更を察知することが難しい。

上記の2点について解決したい。


### 改善観点

- レンダリングされる*Component*の数をチューニング
- useEffectの実行タイミングの調整

### 初期状態（サンプルデモ）

<iframe src="https://codesandbox.io/embed/github/dede-20191130/reactpj01/tree/tuning-add-useref-derv-no-tuned/?fontsize=14&hidenavigation=1&theme=dark"
    style="width:100%; height:500px; border:0; border-radius: 4px; overflow:hidden;"
    title="reactpj01"
    allow="accelerometer; ambient-light-sensor; camera; encrypted-media; geolocation; gyroscope; hid; microphone; midi; payment; usb; vr; xr-spatial-tracking"
    sandbox="allow-forms allow-modals allow-popups allow-presentation allow-same-origin allow-scripts"
></iframe>

初期表示、および*SetterBoxButton*を二回押下した際のログは次のようになる。  
```
render Container          //*** 初期表示ここから
render Box
render EffectBox
render SetterBox
run useDeepCompareEffect
{foo: {…}, bar: {…}}      //*** ここまで
render Container          //*** 一回目ここから
render Box
render EffectBox
render SetterBox          //*** ここまで
render Container          //*** 二回目ここから
render Box
render EffectBox
render SetterBox          //*** ここまで

```

- ボタン押下ごとに*Container Component*およびすべての子コンポーネントが再レンダリングされている。
- *EffectBox*の*useEffect*が一回目のボタン押下時に実行されない。  
（依存関係に空の配列`[]`を渡しているためにそうなる。本来ならボタン押下で*myContext*が変更した際に実行してほしい）

## チューニングの実行

### レンダリング対象の絞り込み

#### i. React.memoを使用

*Container*の子コンポーネントのそれぞれを、  
React.memoを用いてメモ化する。  

引数の値が変更しなければ、再レンダリングを引き起こさないようにする。  

<iframe src="https://codesandbox.io/embed/tuning-usecontext-for-object--add-memo-ror0u?fontsize=14&hidenavigation=1&theme=dark"
     style="width:100%; height:500px; border:0; border-radius: 4px; overflow:hidden;"
     title="tuning-useContext-for-object--add-memo"
     allow="accelerometer; ambient-light-sensor; camera; encrypted-media; geolocation; gyroscope; hid; microphone; midi; payment; usb; vr; xr-spatial-tracking"
     sandbox="allow-forms allow-modals allow-popups allow-presentation allow-same-origin allow-scripts"
   ></iframe>

ログは次のようになる。

```no{hl_lines=["7-12"]}
render Container         //*** 初期表示ここから                
render Box 
render EffectBox
render SetterBox
run useDeepCompareEffect
{foo: {…}, bar: {…}}     //*** ここまで
render Container         //*** 一回目ボタン押下 ここから
render Box
render EffectBox         //*** ここまで
render Container         //*** 二回目ボタン押下 ここから
render Box
render EffectBox         //*** ここまで
  
```

子コンポーネントのうち、  
`SetterBox`は再レンダされなくなった。  

しかし、  
`Box`および`EffectBox`はメモ化しても再レンダされる。  

ボタン押下時に*myContext*が更新され、*myContext.Provider*で渡される値が変化するため、  
`useContext(myContext)`によって*myContext*を使用している*Component*はすべて再レンダされる。  
ここで渡される値はオブジェクトの参照値なので、たとえ内容が等しくても「別のオブジェクト」と判定され、  
レンダリングが実行される。

次のiiで、その挙動を改善したい。


#### ii. useRefを使用してDeep Compareを実行

*Container Component*の`cxt`変数が更新されたあと、  
*useRef*でフックされた`cxtRef`変数と、  
`Deep Equal`（ネストされた各値のすべてが等しいことを保証）によって内容を比較する。  
*myContext.Provider*のvalueに渡されるのは`cxt`ではなく`cxtRef.current`となる。

`Deep Equal`の実装には、外部ライブラリの`fast-deep-equal`を使用する。

<iframe src="https://codesandbox.io/embed/tuning-usecontext-for-object--add-useref-p49b7?fontsize=14&hidenavigation=1&theme=dark"
     style="width:100%; height:500px; border:0; border-radius: 4px; overflow:hidden;"
     title="tuning-useContext-for-object--add-useRef"
     allow="accelerometer; ambient-light-sensor; camera; encrypted-media; geolocation; gyroscope; hid; microphone; midi; payment; usb; vr; xr-spatial-tracking"
     sandbox="allow-forms allow-modals allow-popups allow-presentation allow-same-origin allow-scripts"
   ></iframe>

ログは次のようになる。

```no{hl_lines=["10"]}
render Container             //*** 初期表示ここから
render Box
render EffectBox
render SetterBox
run useDeepCompareEffect
{foo: {…}, bar: {…}}       //*** ここまで
render Container             //*** 一回目ボタン押下 ここから
render Box
render EffectBox             //*** ここまで
render Container             //*** 二回目ボタン押下 ここからここまで
  
```

二回目のボタン押下では、  
*myContext*の内容は変更しないため、  
`Deep Compare`の結果`cxtRef.current`変数の参照は変化しない。  

したがって、`Box`および`EffectBox`の再レンダは引き起こされない。

### useEffectの実行タイミングの調整

*useEffect*の依存関係においてmutableな値を許容するトリックだが、  
検索した結果`useDeepCompareEffect`という便利なフックが公開されていた。  

[こちらの記事の下の方に説明が記載されています。](https://www.benmvp.com/blog/object-array-dependencies-react-useEffect-hook/#option-4---do-it-yourself)

これを用いることによって、  
*myContext*が変更したタイミングのいずれにおいても  
EffectBoxのサイドエフェクト処理が走るようにできる。

<iframe src="https://codesandbox.io/embed/tuning-usecontext-for-object--complete-3nybn?fontsize=14&hidenavigation=1&theme=dark"
     style="width:100%; height:500px; border:0; border-radius: 4px; overflow:hidden;"
     title="tuning-useContext-for-object--complete"
     allow="accelerometer; ambient-light-sensor; camera; encrypted-media; geolocation; gyroscope; hid; microphone; midi; payment; usb; vr; xr-spatial-tracking"
     sandbox="allow-forms allow-modals allow-popups allow-presentation allow-same-origin allow-scripts"
   ></iframe>

ログは次のようになる。

```no{hl_lines=["7-11"]}
render Container             //*** 初期表示ここから
render Box
render EffectBox
render SetterBox
run useDeepCompareEffect
{foo: {…}, bar: {…}}         //*** ここまで
render Container             //*** 一回目ボタン押下 ここから
render Box
render EffectBox             
run useDeepCompareEffect
{foo: {…}, bar: {…}}　　　　  //*** ここまで
render Container             //*** 二回目ボタン押下 ここからここまで
  
```

一回目のボタン押下で*myContext*の内容が変更されるため、  
`EffectBox`内のサイドエフェクト処理が実行され、ログに出力されるようになった。

