---
title: "ご意見送信フォームを作成する①　[TypeScript]"
author: dede-20191130
date: 2021-10-08T11:42:48+09:00
slug: send-contact-form-plain-ts
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
    フレームワークなしのTypeScriptを用いて意見送信フォームを作成した。<br>
    Model-View-Controllerの3モデルを用いて情報・表示・動作を管理した。
{{< /box-with-title >}}

フロントエンド開発の練習として、  
クラウドソーシングサイトなどにある案件の開発要望を調べていた。

そちらに、ご意見送信フォームの構築の要望があったため、  
似たような要望を想定して自分で作成してみることにした。

[<span id="srcURL"><u>ソースコードはこちらにあります（GitHub）。</u></span>](https://github.com/dede-20191130/send-contact-form)

## 開発要望（想定）

- 名前などのプロファイルや意見内容を入力できるフォームを構築する。
- バリデーションチェックし、送信内容にエラーが有った場合に画面に表示する。
- 意見送信があった場合に、メールアドレスに内容を送信する。また、ユーザが送信後に送信内容を確認できる。

## 設計

### シーケンス設計

要望を受けた上での処理の流れをどのようにするかを考え、次のように書き下した。

#### フォーム送信シーケンス

![フォーム送信シーケンス](./send-sequence.png)

#### 送信完了モーダル画面表示シーケンス

![送信完了モーダル画面表示シーケンス](./show-result.png)

■メール送信はメールサーバの設定などが必要で煩雑だったため、  
　メールを送信した体で擬似的にテキストファイルを作成して、ダウンロードできるようにしました。

### クラス設計

#### About

フロントエンドにおいてMVCモデルを適用するイメージで作成する。

アプリの機能の役割を「Model」＋「View」＋「Controller」に分割し、  
それぞれをモジュールとして分離する。

1. Model: ユーザとのインタラクションなどに用いるデータを管理。ビジネスロジックもここに。
2. View： データを表示する機能を担当。
3. Controller： ModelとViewを連携される。

#### クラス一覧

<table>
  <tr>
    <th>所属</th>
    <th>クラス名</th>
    <th>役割</th>
  </tr>
  <tr>
    <td>フォーム</td>
    <td>FormView</td>
    <td>フォームのコンポーネントを参照し、<br>イベントの設定などをおこなう</td>
  </tr>
  <tr>
    <td>フォーム</td>
    <td>FormModel</td>
    <td>フォームの入力値を保持。<br>バリデーション関数を保有。</td>
  </tr>
  <tr>
    <td>フォーム</td>
    <td>FormController</td>
    <td>ViewからModelに入力値を渡したり<br>Viewのイベントを受け取り他のイベントを発火させる。</td>
  </tr>
  <tr>
    <td>送信完了<br>モーダル画面</td>
    <td>ModalView</td>
    <td>モーダル画面のコンポーネントを参照し、<br>イベントの設定などをおこなう</td>
  </tr>
  <tr>
    <td>送信完了<br>モーダル画面</td>
    <td>ModalModel</td>
    <td>送信内容データを保持。<br>ファイル作成ロジックを保持。</td>
  </tr>
</table>

■モーダル画面は処理が簡単のため、Controllerは省略しています。

#### クラス図

![クラス図](./class-diagram.png)

## 実装

### 作成環境

```
node v14.16.1
npm 6.14.12
```
### 依存パッケージ

一部抜粋。

```json
{
  "devDependencies":{
    "babel-loader": "^8.2.2",
      "core-js": "^3.15.2",
      "regenerator-runtime": "^0.13.9",
      "html-webpack-plugin": "^5.3.2",
      "mini-css-extract-plugin": "^2.1.0",
      "css-loader": "^6.2.0",
      "style-loader": "^3.2.1",
      "sass": "^1.35.2",
      "sass-loader": "^12.1.0",
      "typescript": "^4.4.3",
      "webpack": "^5.45.1",
      "webpack-cli": "^4.7.2",
      "moment": "^2.29.1",
   }
}

```

### ディレクトリ構成

一部抜粋。

```
myapp
│  .babelrc
│  .eslintrc.js
│  jest.config.js
│  package-lock.json
│  package.json
│  webpack.config.js
│  
├─node_modules
│          
└─src
    │  index.html
    │  tsconfig.json
    │  
    ├─style
    │      style.scss
    │      _active.scss
    │      _form.scss
    │      _mixin.scss
    │      _modal.scss
    │      _variables.scss
    │      
    └─ts
       │  index.ts
       │  NameSpace.ts
       │  
       ├─form
       │      FormController.ts
       │      FormErrorMessages.ts
       │      FormModel.ts
       │      FormView.ts
       │      
       └─modal
               ModalModel.ts
               ModalView.ts
```

### コード

ソースコードの全体は[こちらから](#srcURL)

#### FormView

```ts

interface IFormViewArg {
  address: HTMLInputElement;
    age: HTMLInputElement;
    errArea: HTMLDivElement;
    form: HTMLFormElement;
    gender: RadioNodeList;
    message: HTMLTextAreaElement;
    name: HTMLInputElement;
    submitBtn: HTMLInputElement;
    modalScreen: HTMLDivElement;

}

export class FormView {
  private address: HTMLInputElement;
    private age: HTMLInputElement;
    public errArea: HTMLDivElement;
    private form: HTMLFormElement;
    private formController: FormController;
    private gender: RadioNodeList;
    private message: HTMLTextAreaElement;
    private name: HTMLInputElement;
    private submitBtn: HTMLInputElement;
    constructor({
      form,
        name,
        gender,
        age,
        address,
        message,
        submitBtn,
        errArea,
        modalScreen
    }: IFormViewArg) {
      this.formController = new FormController({
        formView: this,
            modalScreen: modalScreen,
        });
        this.form = form;
        this.name = name;
        this.gender = gender;
        this.age = age;
        this.address = address;
        this.message = message;
        this.submitBtn = submitBtn;
        this.errArea = errArea;

        this.setSubmitEvt(this.submitBtn);

        for (const elem of [this.name, this.age, this.address, this.message]) {
          this.setInputBoxFocusEvts(elem);
        }

    }

    private setSubmitEvt(elem: HTMLElement) {
      elem.onclick = (ev) => {
        ev.preventDefault();
            this.onSubmit();
        };
    }

    private setInputBoxFocusEvts(elem: HTMLElement) {
      elem.onfocus = function (ev) {
        document
                .querySelector(`label[for="${(ev.currentTarget as HTMLElement)?.id}"]`)
                ?.classList.add("active");
        };
        elem.onblur = function (ev: any) {
          if (ev.currentTarget.value === "") {
            
            document
                    .querySelector(`label[for="${(ev.currentTarget as HTMLElement)?.id}"]`)
                    ?.classList.remove("active");

            }
        };
    }

    private onSubmit() {
      this.formController.onSubmit({
        name: this.name.value,
            gender: this.gender.value,
            age: this.age.value,
            address: this.address.value,
            message: this.message.value,
        });
    }
}

```

`onSubmit`メソッドで`Controller`の`onSubmit`メソッドを呼び出して  
モーダルを表示するためのデータを引き渡す。

`setInputBoxFocusEvts`メソッドによって各input要素にイベントを設定し、  
要素を選択した際にラベルがきれいに移動するようにしている。

#### FormModel

```ts
export interface IFormModelArg {
    name: string;
    gender: string;
    age: string;
    address: string;
    message: string;
};

export class FormModel {
    private address!: string;
    private age!: string;
    private gender!: string;
    private message!: string;
    private name!: string;
    public isvalid: { [method: string]: () => boolean };
    constructor(formData: IFormModelArg) {
        Object.assign(this, formData);
        this.isvalid = {
            name: this.isValidName.bind(this),
            gender: this.isValidGender.bind(this),
            age: this.isValidAge.bind(this),
            address: this.isValidAddress.bind(this),
            message: this.isValidMessage.bind(this),
        };
    }
    private isValidName(): boolean {
        let trimed = this.name.trim();
        return trimed.length > 0 && trimed.length < 21;

    }
    private isValidGender(): boolean {
        const alloweds = [0, 1, 2];
        return alloweds.includes(Number(this.gender));

    }
    private isValidAge(): boolean {
        if (!this.age.trim().length) return false;
        const age = Number(this.age);
        return Number.isInteger(age) && Number(age) > -1;

    }
    private isValidAddress(): boolean {
        let trimed = this.address.trim();
        return trimed.length < 101;

    }
    private isValidMessage(): boolean {
        let trimed = this.message.trim();
        return trimed.length > 0 && trimed.length < 2001;

    }

    public createSerializedData(): string {
        return JSON.stringify({
            name: this.name,
            gender: this.gender,
            age: this.age,
            address: this.address,
            message: this.message,
        });
    }
}
```

`this.isvalid`プロパティが、各入力項目のバリデーション関数を保持する役割を持つ。  
`createSerializedData`メソッドにより、各入力地をJSON形式で他のクラスに渡す。


#### FormController

```ts
interface IFormControllerArg {
    formView: FormView;
    modalScreen: HTMLDivElement;
};

export class FormController {
    private _formModel: FormModel | undefined;
    private _formView: FormView;
    private _modalScreen: HTMLDivElement;
    constructor({
        formView,
        modalScreen
    }: IFormControllerArg) {
        this._formView = formView;
        this._formModel;
        this._modalScreen = modalScreen;
    }
    public onSubmit({
        name,
        gender,
        age,
        address,
        message
    }: IFormModelArg): void {
        this._formView.errArea.innerHTML = "";

        this._formModel = new FormModel({
            name: name,
            gender: gender,
            age: age,
            address: address,
            message: message,
        });

        let errFounds = this.isvalid();
        if (errFounds) {
            this.setError(errFounds);
            this._formView.errArea.scrollIntoView(false);
            return;
        }

        this._modalScreen.dispatchEvent(
            new CustomEvent("show", {
                detail: {
                    serializedData: this._formModel.createSerializedData(),
                },
            })
        );
    }
    private isvalid() {
        let errFounds = [];
        for (const prop of ["name", "gender", "age", "address", "message"]) {
            if (this._formModel && !this._formModel.isvalid[prop]()) errFounds.push(prop);
        }
        return errFounds.length === 0 ? null : errFounds;
    }
    private setError(errFounds: string[]) {
        this._formView.errArea.innerHTML = errFounds.reduce((acc, curr) => {

            return acc + formErrorMessages[curr] + "<br>";
        }, "");
    }
}

```

Viewから入力値を受け取ってModelに渡し、  
Modelよりバリデーション結果を受け取る。  

バリデーションエラー発生時はViewにエラー内容を渡し、  
正常にSubmitされた場合は`CustomEvent`の`[show]`イベントを発火させる。  
イベント発火時、`FormModel`から渡されたJSONデータをモーダル要素に渡す。

#### ModalView

```ts
interface IModalViewArg {
    screen: HTMLDivElement;
    screenCover: HTMLDivElement;
    dlBtn: HTMLButtonElement;
    closeBtn: HTMLDivElement;

}

export class ModalView {
    private screen: HTMLDivElement;
    private screenCover: HTMLDivElement;
    private dlBtn: HTMLButtonElement;
    private closeBtn: HTMLDivElement;
    private modalModel: ModalModel;
    constructor({
        screen,
        screenCover,
        dlBtn,
        closeBtn
    }: IModalViewArg) {
        this.modalModel = new ModalModel();
        this.screen = screen;
        this.screenCover = screenCover;
        this.dlBtn = dlBtn;
        this.closeBtn = closeBtn;

        // avoid custom-event caveat by below
        // https://github.com/microsoft/TypeScript/issues/28357#issuecomment-436484705
        this.screen.addEventListener("show", ((ev: CustomEvent) => {
            this.modalModel.serializedData = ev.detail.serializedData;
            this.screen.hidden = !this.screen.hidden;
            this.screenCover.hidden = !this.screenCover.hidden;
            document.body.classList.add("preventScroll");
        }) as EventListener);
        this.dlBtn.onclick = this.download.bind(this);
        this.closeBtn.onclick = this.close.bind(this);
    }
    private download() {
        const text = this.modalModel.createText();
        const link = document.createElement("a");
        link.download = "受理内容.txt";
        link.href = this.modalModel.createTextBlob(text);

        link.click();

        URL.revokeObjectURL(link.href);
    }
    private close() {
        this.screen.hidden = !this.screen.hidden;
        this.screenCover.hidden = !this.screenCover.hidden;
        document.body.classList.remove("preventScroll");
    }
}

```

`show`イベントが発火されると、  
スクリーンカバー用の`div`要素とModal要素を出現させる。

ダウンロードボタン、☓ボタン押下時のイベントをそれぞれ持つ（`download`, `close`）。  

`download`イベント発火時、  
フォームに入力した値を、受理内容.txtというテキストファイルとしてローカルにダウンロードする。

#### ModalModel

```ts
const textTemplate = `※ご意見フォーム送信フェイク※
下記内容で承りました。

【氏名】$name
【性別】$gender
【年齢】$age歳
【住所】$address
【ご意見内容】
$message

-----------------------

受理日時：$date

`;

export class ModalModel {
    private _serializedData: string = "";
    set serializedData(value: string) {
        this._serializedData = value;
    }
    public createText() {
        const data = JSON.parse(this._serializedData) as IFormModelArg;

        // convert gender:int to string
        data.gender = ["その他", "男性", "女性"][Number(data.gender)];

        let text = textTemplate;
        let key: keyof IFormModelArg;
        for (key in data) {
            if (Object.hasOwnProperty.call(data, key)) {
                text = text.replace("$" + key, data[key]);
            }
        }
        // hack for jest/typescript on testing moment
        // https://newbedev.com/using-rollup-for-angular-2-s-aot-compiler-and-importing-moment-js
        const momentFunc = (moment as any).default ? (moment as any).default : moment;
        text = text.replace("$date", momentFunc().format("YYYY年MM月DD日"));

        return text;
    }
    public createTextBlob(text: string) {
        const blob = new Blob([text], { type: "text/plain" });
        return URL.createObjectURL(blob);
    }
}
```

`createText`メソッドにより、JSONデータをテキストデータに変換する。

## デモ（Vercel）

[Vercelにデプロイしました。](https://send-contact-form.vercel.app/)

## 終わりに

練習として、  
Reactなどのフレームワークは使用せずにプレーンなTypescriptで実装した。

こうしてみると、各種フレームワークがどのような要望や要請をもとに作成されているのか  
（フレームワークを使わないことによる作業のボトルネックはどこにあるか）が少しはわかるような気がする。






