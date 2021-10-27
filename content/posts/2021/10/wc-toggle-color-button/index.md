---
title: "[JavaScript] そうだ、Web Componentsを使おう"
author: dede-20191130
date: 2021-10-27T10:01:42+09:00
slug: wc-toggle-color-button
draft: true
toc: true
featured: false
tags: ["Javascript","フロントエンド"]
categories: ["プログラミング"]
archives:
    - 2021
    - 2021-10
---

## この記事について

{{< box-with-title title="かんたんな概要と結論" >}} 
    <i>Web Components</i>の利点と、利点を表現した簡単なデモを作成した。<br>
    Componentのカプセル化を用いることで、外部に影響されないスタイル、動作を設計することができる。
{{< /box-with-title >}}

JavaScriptの標準APIである[*Web Components*](https://developer.mozilla.org/ja/docs/Web/Web_Components)を学習している際に、  
一度自分でもわかりやすいようにデモを作成してみようと思った。

今回は、画面あるいは指定ボックス内の明るいモード、暗いモードを切り替えるトグルボタンを、  
*Web Components*を用いて作成した。

[<span id="srcURL"><u>こちらからダウンロードできます（GitHub）。</u></span>](https://github.com/dede-20191130/wc-toggle-color-button)

## Web Componentsとは？

### About

>Web Components は、再利用可能なカスタム要素を作成し、ウェブアプリの中で利用するための、一連のテクノロジーです。コードの他の部分から独立した、カプセル化された機能を使って実現します。 [https://developer.mozilla.org/ja/docs/Web/Web_Components](https://developer.mozilla.org/ja/docs/Web/Web_Components)

Reactのような外部ライブラリを用いなくても、  
標準APIのみで、DOMの一部とデザイン、イベントなどを一つのコンポーネントにまとめて使い回すことができる技術。

### サンプル

下記「サンプルを作成する」で作成したサンプル（説明は後述）。  
'🌞'のマークのトグルボタンをクリックすると、画面やボックスの色を切り替える。

<iframe src="https://codesandbox.io/embed/wc-toggle-color-button-iwnvm?fontsize=14&hidenavigation=1&theme=dark"
     style="width:100%; height:500px; border:0; border-radius: 4px; overflow:hidden;"
     title="wc-toggle-color-button"
     allow="accelerometer; ambient-light-sensor; camera; encrypted-media; geolocation; gyroscope; hid; microphone; midi; payment; usb; vr; xr-spatial-tracking"
     sandbox="allow-forms allow-modals allow-popups allow-presentation allow-same-origin allow-scripts"
   ></iframe>

### Pros

- コンポーネントの使い回しを可能とする  
DRYなコーディングを助ける。
- styleをカプセル化することができる 
これは、あるセレクタに該当する要素に対する外部で指定されたstyleが、  
コンポーネント内（正確にはコンポーネント内の*Shadow DOM*）に影響を与えず、  
また、コンポーネント内で記述されたstyleが外部の要素に影響を与えることはないということを意味している。
- 動作のカプセル化をすることができる  
styleと同様に、clickイベントなども独自の挙動を定義することが可能。

### Cons

- コンポーネントの外部と内部とのデータのやり取りが苦手  
*Shadow DOM*を用いてカプセル化する場合は特に制限が強く、  
Reactのように自由気ままにイベント移譲をしたりstate管理したりすることは困難。
- サポートしているブラウザに制限がある  
今後の開発に期待。

## サンプルを作成する（明暗モード切替ボタン）

### About

画面、あるいは指定ボックス内の明るいモード、暗いモードを切り替えるトグルボタンを作成したい。

### 要件

- スタイルやクリック時のイベントを内部で設定したい
- スタイルの一部（ボタンの大きさ）は外部で定義した値を使えるようにする
- 切り替え対象の要素も、外部で定義した値を使えるようにする

### コード

#### toggle-color-button-component.ts

```ts
const style = `
:host {
    --toggle-color-button-criteria-len:var(--toggle-color-button-len,75px);
    width: var(--toggle-color-button-criteria-len);
    height: calc(var(--toggle-color-button-criteria-len) / 1.78);
}
div{
    position: relative;
    width: var(--toggle-color-button-criteria-len);
    height: calc(var(--toggle-color-button-criteria-len) / 1.78);
    margin: auto;
    display: inline-block;
}
input {
    position: absolute;
    left: 0;
    top: 0;
    width: 100%;
    height: 100%;
    margin:0;
    z-index: 5;
    opacity: 0;
    cursor: pointer;
}
label {
    width: 100%;
    height: 100%;
    background: #ffeb3b;
    position: relative;
    display: inline-block;
    border-radius: calc(var(--toggle-color-button-criteria-len)/1.63);
    transition: 0.4s;
    box-sizing: border-box;
}
label:after {
    content: '🌞';
    position: absolute;
    left: 0;
    top: 0;
    z-index: 2;
    transition: 0.4s;
    font-size: calc(var(--toggle-color-button-criteria-len)/2.5);
    line-height: calc(var(--toggle-color-button-criteria-len)/1.70);
}
input:checked + label {
    background-color: #3d00a9;
}
input:checked + label:after {
    content: '🌙';
    left: calc(var(--toggle-color-button-criteria-len) / 2.14);
}
`;

const template = `
<div>
    <input id="toggle" type='checkbox' />
    <label for="toggle" />
</div>
`;

const tmpl = document.createElement("template");
tmpl.innerHTML = `<style>${style}</style>${template}`;

customElements.define("toggle-color-button", class extends HTMLElement {
    connectedCallback() {
        // shadowDOMの設定
        const shadowRoot = this.attachShadow({ mode: "closed" });
        shadowRoot.append(tmpl.content.cloneNode(true));
        // トグルする対象の要素
        const toggledElem = document.querySelector<HTMLElement>(this.dataset.toggled || "html");
        if (!toggledElem) return;
        // ボタンクリック時にトグル
        shadowRoot.querySelector("input")?.addEventListener("click", () => {
            toggledElem.dataset.mode =
                toggledElem.dataset.mode !== 'dark' ?
                    'dark'
                    : 'light';
        })
    }
});
```

`template`要素にコンポーネントのDOM構造を記述する。  
内部で使われるstyleも文字列として記述。

なお、ここではstyleを文字列としてベタ書きしたけれども、  
Sass用のライブラリやCSS-in-JSのライブラリを用いてうまいことトランスパイルしてあげれば  
保守性を高めることができるだろう（*Web Componetns*内部に記述するstyleなどそこまで多くはないので、だいたいは通常のCSSで事足りるだろうけど）

*customElements.define*メソッド実行によって  
スクリプトを読み込んだHTML側で`<toggle-color-button></toggle-color-button>`の形でカスタムタグを使用することができるようになる。

#### index.html

```html
<!DOCTYPE html>
<html lang="en">

<head>
    <meta charset="UTF-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Document</title>
    <script src="./dist/toggle-color-button-component.js"></script>
    <style type="text/css">
        #box-one {
            --toggle-color-button-len: 50px;
            margin: 60px auto;
            max-width: 400px;
            border: solid 2px;
        }

        #box-one[data-mode="dark"] {
            background-color: darkblue;
            color: lightblue;
        }

        #box-two {
            --toggle-color-button-len: 40px;
            margin: 60px auto;
            max-width: 600px;
            border: solid 2px;
            text-align: center;
        }

        #box-two input {
            display: none;
        }

        #box-two label:before {
            font-family: FontAwesome;
            display: inline-block;
            content: "□";
            color: blue;
            letter-spacing: 10px;

        }

        #box-two input:checked+label:before {
            content: "✔";
        }

        html[data-mode="dark"] * {
            background-color: darkslategrey;
            color: rgb(255, 230, 0);

        }

        label,
        input {
            cursor: pointer;
        }
    </style>
</head>

<body>
    <div id="box-one">
        <p>ボックス内のダークモード切り替えボタン：<toggle-color-button data-toggled="#box-one"></toggle-color-button>
        </p>

    </div>
    <div id="box-two">
        <div>
            画面全体のダークモード切り替えボタン：<toggle-color-button></toggle-color-button>
        </div>
        <p><input id="switchA" type="checkbox" /><label for="switchA">スイッチA</label></p>
        <p><input id="switchB" type="checkbox" /><label for="switchB">スイッチB</label></p>
        <p><input id="switchC" type="checkbox" /><label for="switchC">スイッチC</label></p>
        <p><input id="switchD" type="checkbox" /><label for="switchD">スイッチD</label></p>
        <hr>
        <div>
            <p>スイッチ稼働状態</p>
            <p>スイッチA：<span id="result-switchA">OFF</span></p>
            <p>スイッチB：<span id="result-switchB">OFF</span></p>
            <p>スイッチC：<span id="result-switchC">OFF</span></p>
            <p>スイッチD：<span id="result-switchD">OFF</span></p>
        </div>

    </div>
    <script>
        for (const input of document.querySelectorAll("#box-two input")) {
            input.addEventListener("click", (event) => {
                const state = event.currentTarget.checked;
                document.querySelector("#result-" + event.target.id).textContent = state ? "ON" : "OFF";
            })
        }

    </script>
</body>

</html>
```

##### 切り替え対象の要素の指定

`box-one`要素では、  
`data-toggled`属性として`#box-one`を指定することで、  
明暗モードを切り替える対象を`box-one`内に限定している。

これを指定しない場合、画面全体（ルート要素=*html*要素）が切り替えられることになる

##### styleの独立

`box-two`要素内で、  
チェックボックスやlabelなどに色や大きさなどのstyleを指定しているが、  
`box-two`内部にあるはずの*Web Components*には影響しない。

また、*Web Components*でグローバルに宣言したチェックボックスなどのstyleも、  
`box-two`内のほかのチェックボックスに影響しない。

##### 動作の独立

`box-two`内のチェックボックスは、  
それぞれ対応する稼働状態表示欄にON/OFFを報告するイベントを設定している。

もし*Shadow DOM*によるカプセル化が行われていなければ、  
*Web Components*内部のチェックボックスもまた同様のイベントが設定され、対応する表示欄がないためにエラーでスクリプトが落ちる。

しかし、*Shadow DOM*のおかげで、ここではコンポーネントはそのイベントが設定されず、エラーが起こらない。

### デモ

[CodeSandBox](https://codesandbox.io/s/wc-toggle-color-button-iwnvm)

<iframe src="https://codesandbox.io/embed/wc-toggle-color-button-iwnvm?fontsize=14&hidenavigation=1&theme=dark"
    style="width:100%; height:500px; border:0; border-radius: 4px; overflow:hidden;"
    title="wc-toggle-color-button"
    allow="accelerometer; ambient-light-sensor; camera; encrypted-media; geolocation; gyroscope; hid; microphone; midi; payment; usb; vr; xr-spatial-tracking"
    sandbox="allow-forms allow-modals allow-popups allow-presentation allow-same-origin allow-scripts"
></iframe>

