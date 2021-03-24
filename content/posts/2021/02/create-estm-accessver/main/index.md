---
title: "[Access VBA] 見積書作成ツール（Accessバージョン）を作成した"
author: dede-20191130
date: 2021-02-28T00:34:49+09:00
slug: create-estm-accessver
draft: true
toc: true
featured: false
categories: ["アプリケーション","プログラミング"]
tags: ["Access","VBA","自作","ツール"]
vba_taxo: help_office
archives:
    - 2021
    - 2021-02
---

## この記事について

以前、  
こちらの記事で紹介したように、  
Excelベースの見積書作成ツールを作成した。  

{{< box-with-title title="記事：" >}} 
    {{< page-titled-link page="create-estm" >}}
{{< /box-with-title >}}

今回、ほぼ同機能を持つツールを、  
Accessベースに移植した。

理由としては、  
- 一対多のリレーションシップを持つデータを管理、抽出するのは  
Accessのほうが遥かに容易である。
- GUIの作成のための機能に関して、Accessがよりリッチである。
- Accessでのツール作成技術のさらなる向上のため。

がある。

## 概要

各画面で、  
あらかじめ設定したマスタデータを選択し、  
テンプレート見積書に設定したデータを挿入する。  

Excelブック形式で  
見積書を出力する。

出力した見積書に使用したデータを保持し、  
再利用できるようにする。


## 動作環境
- Windows
- 2013以上のOfficeソフトが動作する環境であれば可。  
2010以下のOfficeでも動作する可能性はあります。  
（そこまでの下位Ver.互換性に需要があるかどうかは不明なため、検証はしていません）。

## ツール外観（各画面紹介）

{{< page-titled-link page="../slides/tool-screen"  title="Click To Show Slide Page" relative=true  >}}
