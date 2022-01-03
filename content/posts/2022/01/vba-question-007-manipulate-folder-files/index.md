---
title: "[教えて！VBA] 第7回 フォルダを開く（＋ファイルを操作する）にはどうすればいいの？？"
author: dede-20191130
date: 2022-01-02T13:07:25+09:00
slug: vba-question-007-manipulate-folder-files
draft: false
toc: true
featured: false
tags: ["VBA","Excel"]
categories: ["プログラミング"]
archives:
    - 2022
    - 2022-01
---



![](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1641189331/learnerBlog/vba-question-007-manipulate-folder-files/manipulate-folder-files_m8m4zz.png)

## この記事について

{{< box-with-title title="かんたんな概要と結論" >}} 
    フォルダを開くには、<br/>エクスプローラを起動して開く、ダイアログボックスで開くの二通りが存在します。<br/><br/>
    また、操作したいファイルが有るフォルダが固定であれば、<br/>APIを用いてバックグラウンドでファイルを開いて操作することができます。
{{< /box-with-title >}}

こんにちは、dedeです。

この記事では、  
VBAマクロに関する質問のうち、  
皆が疑問に思っているトピックについて解説いたします。

今回は、
{{< colored-span color="#fb9700" >}}VBAマクロからフォルダを開く（＋α ファイルを操作する）方法{{< /colored-span >}}  
を解説いたします。

※この記事は、Office VBAマクロのうち  
Excel VBAマクロに関するトピックです。

レベル：<b>中級者向け</b>

## やりたいこと

特定フォルダの内容をチェックしたり、  
CSVなどの外部ファイルの内容を取り込んだりすることは  
VBAマクロにおいてしばしば要求される処理の一つです。

マクロで「フォルダを開く」という動作にはいくつか種類があり、  
1. エクスプローラ（ファイル閲覧、管理するアプリケーション）上でフォルダを開く
2. ダイアログボックス画面を開き、ユーザにファイル・フォルダを選択させる
3. 組み込み関数やAPIを用いてフォルダ内を探索する
などを、目的に合わせて採用することになると思われます。

ただフォルダを開いて内容を視認することが目的であれば1. を、  
利用者が操作するファイルを選ぶことができることを目的とするのであれば2. を、  
すでに操作するファイルの置かれたフォルダの場所が決まっているのであれば3. を採用するのが良いでしょう。

## フォルダを開くだけの場合

フォルダを開くには、
1. ハイパーリンクの仕組みを利用する方法
2. 外部プログラムを実行する方法（Shell）
の二種類が存在します。

![フォルダを開くだけの場合](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1641189331/learnerBlog/vba-question-007-manipulate-folder-files/manipulate-folder-files2_bdhtna.png)

### ハイパーリンクの仕組みを利用する方法

#### ABOUT

`Workbook`オブジェクトの`FollowHyperlink`メソッドを利用します。

```vb
Workbook.FollowHyperlink(引数)
'// メソッドの詳細はこちらを参照 https://vbabeginner.net/open-web-page-without-hyperlink/
```

このメソッドは、  
Excelシート上のハイパーリンク押下時の挙動と同様の動作をしてフォルダを開きます。

ご存知かもしれませんが、  
Excelシートにハイパーリンクを設定して、WebサイトのURLやブック内のセル番地ではなく  
フォルダのパス（C:\temp~など）を指定すると、  
リンクを押下したときにエクスプローラが起動し、指定したフォルダが開かれます。

その仕組みを利用します。

#### コード

次のコードを実行すると、  
引数`Address`で指定したパスのフォルダを、  
エクスプローラで開きます。

```vb

Sub ハイパーリンクの仕組みを利用する関数()
    Dim path As String
    path = "C:\temp\"
    Application.ThisWorkbook.FollowHyperlink Address:=path
End Sub
```

### 外部プログラムを実行する方法（Shell）
#### ABOUT

Shell関数というVBAにもともと備わっている関数を利用します。

Shell関数とは、  
外部の実行可能プログラム（アプリケーション）を指定して実行する事ができる機能を持っている関数です。

起動可能なアプリケーションは、  
エクスプローラ、メモ帳、拡張子exeの各種ファイルなど多岐にわたります。

今回は、  
Shell関数を利用してエクスプローラからフォルダを開きます。
#### コード

次のように書きます。

```vb

Sub Shell関数を利用するパターン()
    Call Shell("explorer ""C:\temp\""", vbNormalFocus)
    '// 詳細な引数の指定方法はこちらを参照 https://vbabeginner.net/shell/
End Sub

```

最初の引数には、  
利用するアプリケーションパス（今回は`explorer`。環境変数でパス解決されるため、explorerの文字だけで問題なし）と、  
その後にフォルダパス名を記載する文字列を指定します。

その後、2つ目の引数で、  
起動時のエクスプローラの開き方（最小化、最大化、非表示など）を指定します。  
今回は`vbNormalFocus`として、前回起動時のウィンドウ状態に準拠させます。

## フォルダの中のファイルを操作する場合

フォルダを開くだけでなく、  
フォルダ内のファイルを操作したい場合は、  
上述の
2. ダイアログボックス画面を開き、ユーザにファイル・フォルダを選択させる
3. 組み込み関数やAPIを用いてフォルダ内を探索する
の方法を利用します。


### ダイアログボックスから手動で選択する場合
#### ABOUT

![ダイアログボックスから手動で選択する場合](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1641189331/learnerBlog/vba-question-007-manipulate-folder-files/manipulate-folder-files3_th9e3x.png)

上の図のように、  
マクロからダイアログボックスを起動し、  
手動で選択したファイルを取り扱います。

#### コード

メイン関数である「ダイアログボックス経由でファイル情報出力」から、  
機能別に分けられたサブ関数２種類を呼び出しています。

```vb
Sub ダイアログボックス経由でファイル情報出力()
    Dim filePath As String
    
    filePath = サブ関数_ダイアログボックス表示してファイルパス取得()
    
    If filePath <> "" Then
        サブ関数_ファイルの情報をシートに出力 (filePath)
    End If
    
End Sub

Function サブ関数_ダイアログボックス表示してファイルパス取得() As String
    With Application.FileDialog(msoFileDialogFilePicker)
        '//フィルタ設定 テキストファイルだけ選択可能に制限する
        .Filters.Clear
        .Filters.Add "テキストファイル", "*.txt"
        
        '//複数ファイル選択不可能に設定
        .AllowMultiSelect = False
        .Title = "ファイル選択"
        
        '//ファイル選択した場合に「Show」がTrueになる
        '//ファイル選択せず閉じた場合はFalse
        If .Show = True Then
            サブ関数_ダイアログボックス表示してファイルパス取得 = .SelectedItems(1)
        End If
    End With
End Function

Sub サブ関数_ファイルの情報をシートに出力(ByVal filePath As String)
    Dim fso As Object
    Dim myFile As Object
    '//ファイル取り扱う機能を持つオブジェクトを取得（FileSystemObject）
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    '//パスで指定されたファイルを取得
    Set myFile = fso.GetFile(filePath)
    
    '//A1セルにファイル作成日時を出力
    ActiveSheet.Range("A1").Value = myFile.DateCreated
    
    '//A2セルにテキストファイルの一行目の文字列を出力
    With myFile.OpenAsTextStream
        ActiveSheet.Range("A2").Value = .ReadLine
        .Close
    End With
End Sub

```

#### デモ

上記コードを実行すると、  
ダイアログボックスが表示され、ファイルを選択できます。

その後、  
シートのA1セルにファイルの作成日時が、  
A2セルにファイルの一行目の内容が出力されます。

![マクロ実行前](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1641189330/learnerBlog/vba-question-007-manipulate-folder-files/%E3%82%B9%E3%82%AF%E3%83%AA%E3%83%BC%E3%83%B3%E3%82%B7%E3%83%A7%E3%83%83%E3%83%88_2022-01-03_125410_ij4ul3.png)

![マクロ実行後](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1641189330/learnerBlog/vba-question-007-manipulate-folder-files/%E3%82%B9%E3%82%AF%E3%83%AA%E3%83%BC%E3%83%B3%E3%82%B7%E3%83%A7%E3%83%83%E3%83%88_2022-01-03_125443_i3wzcm.png)


### 組み込み関数やAPIを利用する場合
#### ABOUT

ダイアログボックスを利用しない、  
つまり探索するフォルダがすでに決まっているならば、  
VBA組み込み関数（`Dir`）やファイルシステムAPI（`FileSystemObject`）を利用して、  
バックグラウンドですみやかにファイルの操作を行うことができます。

`Dir`（および他のVBA組み込み関数）よりも`FileSystemObject`のほうが  
直感的で高性能のため、  
以下では`FileSystemObject`を利用してフォルダ内のファイルの探索をするデモコードを説明します。

#### コード

```vb

Sub 組み込み関数やAPIを利用するパターン()
    Dim fso As Object
    Dim folderPath As String
    Dim myFolder As Object
    Dim myFile As Object
    Dim i As Long
    
    '//ファイル取り扱う機能を持つオブジェクトを取得（FileSystemObject）
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    '//フォルダパス（固定：ユーザには選択させない）
    folderPath = "C:\temp\"
    
    '//フォルダの取得
    Set myFolder = fso.GetFolder(folderPath)
    
    '//フォルダ中のテキストファイル（拡張子txt）のファイルのみ対象として
    '//一行目の文字列データをシートに出力
    i = 1
    For Each myFile In myFolder.Files
        '//テキストファイルかどうかをチェック
        If fso.GetExtensionName(myFile.path) = "txt" Then
            '//セルにテキストファイルの一行目の文字列を出力
            With myFile.OpenAsTextStream
                ActiveSheet.Cells(i, 1).Value = .ReadLine
            End With
            i = i + 1
        End If
    Next
    
End Sub
```
#### デモ

tempフォルダに次のようにファイルが格納されているとします。

```ps

Mode                 LastWriteTime         Length Name
----                 -------------         ------ ----
-a----        2022/01/03     14:42             76 another-text01.csv
-a----        2022/01/03     14:42             76 another-text02.json
-a----        2022/01/03     13:38             76 mytext01.txt
-a----        2022/01/03     14:20             76 mytext02.txt
-a----        2022/01/03     14:42             76 mytext03.txt

```

コードを実行すると、  
mytext（拡張子txtのファイル）のみを対象として、  
シートにファイルの内容を出力します。

![マクロ実行前](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1641189331/learnerBlog/vba-question-007-manipulate-folder-files/%E3%82%B9%E3%82%AF%E3%83%AA%E3%83%BC%E3%83%B3%E3%82%B7%E3%83%A7%E3%83%83%E3%83%88_2022-01-03_144848_p2pyrv.png)

![マクロ実行後](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1641189331/learnerBlog/vba-question-007-manipulate-folder-files/%E3%82%B9%E3%82%AF%E3%83%AA%E3%83%BC%E3%83%B3%E3%82%B7%E3%83%A7%E3%83%83%E3%83%88_2022-01-03_144906_x0vyqr.png)
## 終わりに

フォルダの取り扱いには様々あります。

実現したい機能に合わせて、  
適切な方法を選択できるようになると、  
マクロの表現の幅が広がるでしょう。


