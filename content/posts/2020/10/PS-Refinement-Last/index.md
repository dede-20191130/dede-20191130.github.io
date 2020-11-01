---
title: "[Powershell] Windowsで更新日時が〇〇以降のファイルのパスの一覧を取得するには"
author: dede-20191130
date: 2020-10-30T00:00:04+09:00
slug: PS-Refinement-Last
draft: false 
toc: true
tags: ['Powershell']
categories: ['課題解決']
---

## この記事について
普段と別のPC（Windows）での作業をする機会があった。  
作業後にもとのPCに持っていく必要のある、差分あり（作業による変更あり）のファイルを選別する必要があり、  
作業フォルダにおいて、更新日時が本日の9:00以降であるファイルを絞りこむコマンドが欲しかった。  

Git等のバージョン管理アプリがインストールされているPCならば、  
変更内容を適宜コミットしておけばアプリが自動的にうまい具合にやってくれるので、  
このようなコマンドは必要ないのだけれど。  

## 使用環境
PSVersion                      5.1

## コマンド

コンソールに出力
```Powershell
ls -r  -File | ?{$_.LastWriteTime -gt [Datetime]"2020/10/27 9:00:00"} | select   FullName
```
ファイルに出力して見やすくする
```Powershell
ls -r  -File | ?{$_.LastWriteTime -gt [Datetime]"2020/10/27 9:00:00"} | select   FullName | ft  -A   > "C:\temp\output.txt"
```

エイリアス無しコマンドVer
```Powershell
Get-ChildItem -Recurse  -File | Where-Object{$_.LastWriteTime -gt [Datetime]"2020/10/27 18:00:00"} | Select-Object   FullName | Format-Table  -AutoSize   > "C:\temp\output.txt"
```

- Get-ChildItem -Recurse  -File <- 再帰的にファイルのみの情報一覧を取得
- $_.LastWriteTime -gt [Datetime]"2020/10/27 18:00:00" <- 日付文字列をDatetime型にキャストしてファイル情報と比較