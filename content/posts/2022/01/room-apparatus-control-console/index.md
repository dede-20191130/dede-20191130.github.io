---
title: "[ReactJs] 居室内機器管理コンソール画面を作成してみました。"
author: dede-20191130
date: 2022-01-23T15:29:27+09:00
slug: room-apparatus-control-console
draft: false
toc: true
featured: false
tags: ["TypeScript","React","フロントエンド"]
categories: ["プログラミング"]
archives:
    - 2022
    - 2022-01
---


![](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1642950405/learnerBlog/room-apparatus-control-console/room-apparatus-control-console1_marzwp.png)

## この記事について

{{< box-with-title title="かんたんな概要と結論" >}} 
    居室内機器の管理を行うコンソール画面をReactで作成しました。<br/>
    ソースコードは<a href="https://github.com/dede-20191130/room-apparatus-control-console/">こちら</a>にあります。<br/>
    デプロイ先は<a href="https://room-apparatus-control-console.vercel.app/top">こちら</a>です。
{{< /box-with-title >}}

こんにちは、dedeです。

最近、  
フロントエンド開発の練習として、  
クラウドソーシングサイトなどにある案件の開発要望を見ています。

そのときに、居室内機器管理コンソール画面の構築の要望を見かけたため、  
似たような要望を想定して自分で作成してみることにしました。

想定した要件や機能、  
画面別の役割や実装内容について、  
ご説明します。

[<span id="srcURL"><u>ソースコードはこちらにあります（GitHub）。</u></span>](https://github.com/dede-20191130/room-apparatus-control-console/)

## レスポンシブ対応について

モバイル、タブレット機器、PCモニタのどの機器からの閲覧でも対応しています。  
レスポンシブに関するブレークポイントは次のようになっています。

|機器|幅|
|:----|:----|
|モバイル|~520px|
|タブレット|520px~960px|
|PC|960px~|



## 要件


### 概要

Web画面のUI操作により、居室内の機器が制御されます。  
居室内の機器の状態を画面にリアルタイムで表示し、定期的にAPIから情報を取得します。
### 画面構成

共通部品は次の通りです。

- ヘッダー
- サイドバー

個別ページは、  
トップページ、  
各居室の機器運転状態管理ページがあります。

### API

本来ならば別途サーバ用意して  
仮想的に居室内機器を運転させ、その運転状況を管理し、リクエストに対してレスポンスさせるべきだと思っています。

しかしながら、今回はブラウザサイドメインの開発としているため、  
擬似的にサーバにfetchしてデータを取得するAPIモックを構築しました。  
（実際にはサーバは用意せず、ブラウザの`WebStorage`とデータをやり取りするだけです）。
## 機能

- 居室内機器の運転データを表示。定期的に情報を最新化
- 居室内機器の運転を開始/停止
- 居室内機器の運転設定を変更
- 全居室の機器の正常運転/エラー状況を一覧化
- エラー発生機器の復旧をブラウザから実行可能

## 画面
### 共通フレーム

ヘッダーおよびサイドバーがあります。

ヘッダーには運転状態インジケータランプ、および試験エラー発生ボタンがあります。

![共通フレーム](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1642950405/learnerBlog/room-apparatus-control-console/room-apparatus-control-console2_bydw9f.png)

### トップページ

各室運転状況を表示するグリッドボックスが表示されます。  
最大9室まで対応しています。

正常時は緑色、エラー時は赤色として状況が表示されます。  

![トップページ（PC）](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1642950405/learnerBlog/room-apparatus-control-console/room-apparatus-control-console3_ehzcj2.png)

モバイル機器での表示は下のようになります。

![トップページ（モバイル）](https://res.cloudinary.com/ddxhi1rnh/image/upload/c_scale,w_450/v1642950404/learnerBlog/room-apparatus-control-console/room-apparatus-control-console.vercel.app_room101_index_iPhone_12_Pro_qyawd4.png)

### 居室機器インデックスページ

インデックスページでは、  
一枚の居室画像の中に、  
機器のアイコンをそれぞれ表示し、  
状況が早わかりできます（緑色：正常。赤色：エラー）。

![インデックスページ（PC）](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1642950406/learnerBlog/room-apparatus-control-console/room-apparatus-control-console4_slwoi7.png)

モバイル機器での表示は下のようになります。

![インデックスページ（モバイル）](https://res.cloudinary.com/ddxhi1rnh/image/upload/c_scale,w_450/v1642950404/learnerBlog/room-apparatus-control-console/room-apparatus-control-console.vercel.app_room101_index_iPhone_12_Pro_2_patvng.png)
### 個別機器管理ページ

個別機器管理ページでは、  
対象機器の設定項目ごとの設定値および現在値を一覧できます。

設定変更する場合は  
入力フォーム上に入力します。

また、  
エラー発生時には、  
エラーの内容表示、および復旧の手続きを行うことができます。

![個別機器管理ページ（PC）](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1642950405/learnerBlog/room-apparatus-control-console/room-apparatus-control-console5_vgpb3g.png)

モバイル機器での表示は下のようになります。

![個別機器管理ページ（モバイル）](https://res.cloudinary.com/ddxhi1rnh/image/upload/c_scale,w_450/v1642950404/learnerBlog/room-apparatus-control-console/room-apparatus-control-console.vercel.app_room101_index_iPhone_12_Pro_3_xtxazm.png)

## 実装
### 作成環境

```
node v14.16.1
npm 6.14.12
create-react-app 
```

### 依存パッケージ

```json
{
    "dependencies": {
    "@types/node": "^12.20.36",
    "@types/react": "^17.0.33",
    "@types/react-dom": "^17.0.10",
    "@types/react-router-dom": "^5.3.2",
    "@types/styled-components": "^5.1.18",
    "fast-deep-equal": "^3.1.3",
    "react": "^17.0.2",
    "react-dom": "^17.0.2",
    "react-hook-form": "^7.19.5",
    "react-router-dom": "^6.0.2",
    "react-scripts": "4.0.3",
    "styled-components": "^5.3.3",
    "typescript": "^4.4.4",
    }
}

```
### 内容

画面構築のライブラリは*ReactJs*により行っています。

ページ遷移は*React Router*、スタイリングはCSS in JS（*Styled Components*）です。

上記でも触れましたが、機器情報取得などのAPIに関しては、  
簡単のためにAPIモックっぽいものを作ってすべてブラウザで完結するようにしています。
## デモ（Vercelデプロイ）

[Vercelにデプロイしました。](https://room-apparatus-control-console.vercel.app/top)
## 終わりに

*Router*のルーティング設定やレスポンシブなデザイニングなど、  
様々に小課題が現れては対応が必要だったため、  
非常に勉強になりました。

安全なプログラミングとハマりを防ぐために、  
TypeScriptの型システムにもさらに精通する必要があると感じています。