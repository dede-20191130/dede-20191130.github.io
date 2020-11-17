---
title: "[Hugo] Hugoで生成されるデプロイ対象ファイルのコミットまでをスクリプトで一括処理にした"
author: dede-20191130
date: 2020-11-17T10:00:34+09:00
slug: Script-Hugo-Commit
draft: false
toc: true
featured: false
tags: ['Hugo','Git','スクリプト','シェルスクリプト']
categories: ['プログラミング']
archives:
    - 2020
    - 2020-11
---

## この記事について

### Hugoによるブログのデプロイまでの流れ

このブログはHugoという静的サイトジェネレータで管理している。  
デプロイ先は、  
GitHub Pagesという、GitHubに対して静的なウェブサイトをホスティングできるサービス。

大まかにはこのような流れ。

<ol style="list-style-type: lower-roman">
  <li>Hugoの記事ファイル等を作成・編集する。</li>
  <li>hugoコマンドを用いてhtmlファイル等を自動生成する。</li>
  <li>gitでコミットする。</li>
  <li>GitHubのリポジトリにプッシュして<br>ブログを公開する。</li>
</ol>

コマンドの実行やファイルの管理のターミナルは  
<i>git bash</i>を使用している。

### コミットメッセージは必要か？

このとき、  
hugoコマンドで生成するファイルに  
いちいち丁寧なコミットメッセージをつけたりする必要が無いんじゃないかと思いはじめた。  
（自動生成されたファイルの整理や変更、ましてやタイプミスやバグ修正などは全く必要ないため。それらはJavascriptファイルや記事ファイルのほうで対応する）

一人で運用する文には、  
コミットメッセージはすべて「Hugo」で構わないように思える  
（多人数で運用するウェブサイトはちょっと工夫がいるかも知れないけど）。

よって、Hugoのコマンド実行の準備からコミットまでの一連の流れを一括処理できる  
シェルスクリプトを作成した。

## 前提

記事を作成・編集し、hugoコマンドを実行する前

## 作成環境

- Windows10
- git version 2.29.2.windows.2  
- Hugo Static Site Generator v0.76.0/extended windows/amd64


## コードとその内容
### コード

git bashの環境設定ファイル（.bashrc あるいは.bash_profile）に  
下記のように記入する。  

```shell
# エイリアス
alias ga='git add '
alias gc='git commit'

# gitのrootディレクトリ位置に戻る
function groot() {
  if git rev-parse --is-inside-work-tree > /dev/null 2>&1; then
    cd `pwd`/`git rev-parse --show-cdup`
  fi
}

# コミットまでの一括処理
function hugo_commit(){
  groot
  hugo
  ga ./docs
  gc -m 'hugo'
}
```

### groot関数

hugoコマンドを実行するためには、  
一度git管理ディレクトリツリーのrootディレクトリに戻らなければならない。  

そのためgroot関数を実行し、戻る。

[こちらの記事を参考にしました（LINK）](https://qiita.com/ponko2/items/d5f45b2cf2326100cdbc)

### GitHub Pagesのプロジェクトページについて

```shell
ga ./docs
```


GitHub Pagesの静的ウェブサイトデプロイ方法には  
ユーザーページとプロジェクトページの二種類が存在する（私が採用しているのは後者）。  

プロジェクトページでは、設定したブランチの直下のdocsディレクトリを公開対象とするので、  
Hugoのconfigファイルでdocsディレクトリにhtmlファイル等を出力するようにしている。  

そのため、docsディレクトリ配下のみgit addすれば良い。

## 使用法

記事を書き終わって記事ファイル自身のコミットが済んだ後、  
デプロイしたいタイミングでhugo_commitを実行する。

コミット後、リモートリポジトリにプッシュする。
