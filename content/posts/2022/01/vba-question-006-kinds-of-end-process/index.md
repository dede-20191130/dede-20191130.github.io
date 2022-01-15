---
title: "[教えて！VBA] 第6回 エクセルVBAマクロで、マクロを終了するにはどうすればいいの？？"
author: dede-20191130
date: 2022-01-01T08:13:06+09:00
slug: vba-question-006-kinds-of-end-process
draft: false
toc: true
featured: false
tags: ["VBA","Excel"]
categories: ["プログラミング"]
vba_taxo: vbaq
archives:
    - 2022
    - 2022-01
---


![](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1641001541/learnerBlog/vba-question-006-kinds-of-end-process/kinds-of-end-process_h8dftp.png)

## この記事について

{{< box-with-title title="かんたんな概要と結論" >}} 
    マクロを終了するには、①関数から抜け出す、②マクロ全体を終了する、③Excelを閉じるの3通りが存在します<br>
    （解釈次第で他にも存在するでしょう）。<br><br>
    副作用として起こる結果も考慮して、最適な方法を選択すると良いでしょう。
{{< /box-with-title >}}

こんにちは、dedeです。

この記事では、  
VBAマクロに関する質問のうち、  
皆が疑問に思っているトピックについて解説いたします。

今回は、
{{< colored-span color="#fb9700" >}}VBAコードからマクロを終了する方法{{< /colored-span >}}  
を解説いたします。

※この記事は、Office VBAマクロのうち  
Excel VBAマクロに関するトピックです。

レベル：<b>初級者向け</b>

## マクロを終了するとは？（それぞれの要望の違い）

「マクロを終了」という手続きは、  
表現したい処理別に、様々に解釈できると思いますが、  
大きく分けて次の３つに分類されると思っています（他にも色々とあるでしょうが）。

1. 現在の関数から抜け出す。
2. マクロ全体の実行を終了する
3. Excelアプリケーションを閉じる。

下記では、  
１～３のそれぞれについて  
どのようなことが起こっているかの説明と書き方、副作用について  
記載します。

## それぞれの方法についての説明と書き方
### Exit（関数の途中で抜け出す）
#### ABOUT

![Exit（関数の途中で抜け出す）](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1641001541/learnerBlog/vba-question-006-kinds-of-end-process/kinds-of-end-process2_uc7xwg.png)

関数の内部で`Exit Sub`、もしくは`Exit Function`（プロシージャの種類ごとに決定）を呼び出すことで、  
現在実行中の関数から抜け出すことができます。

上の図のように、  
親関数から呼び出した子関数において`Exit Sub`を呼び出すことで、  
子関数のそれ以降の処理を無視して、  
親関数における子関数呼び出し行以降の処理を再開することができます。

#### コード



```vb
Sub 親関数()
    Debug.Print "親関数 " & 1
    Call 子関数
    Debug.Print "親関数 " & 2
End Sub

Sub 子関数()
    Debug.Print "子関数 " & 1
    Exit Sub
    Debug.Print "子関数 " & 2
    Debug.Print "子関数 " & 3
End Sub
```

親関数を実行すると、  
イミディエイトウィンドウにログが表示されます。

```
親関数 1
子関数 1
親関数 2
```

関数の途中で`exit`しているため、  
「子関数 2」の行以降は無視されています。

その後、親関数の「親関数 2」以降の行が順次実行されています。

### End（マクロ全体を実行終了）
#### ABOUT

![End（マクロ全体を実行終了）](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1641001541/learnerBlog/vba-question-006-kinds-of-end-process/kinds-of-end-process3_mnml20.png)

関数の内部で`End`ステートメントを呼び出すことで、  
VBAマクロ全体を終了します。

上の図のように、  
親関数から呼び出した子関数において`End`を呼び出すことで、  
子関数のそれ以降の処理を無視して、  
親関数の処理も無視してマクロ自体を終了します。

※※後述するように、  
この方法にはいくつか副作用があるため、  
注意が必要です。

#### コード

```vb
Sub 親関数()
    Debug.Print "親関数 " & 1
    Call 子関数
    Debug.Print "親関数 " & 2
End Sub

Sub 子関数()
    Debug.Print "子関数 " & 1
    End
    Debug.Print "子関数 " & 2
    Debug.Print "子関数 " & 3
End Sub
```

親関数を実行すると、  
イミディエイトウィンドウにログが表示されます。

```
親関数 1
子関数 1
```

`End`を宣言した後のすべての行は実行されなくなります。  
子関数のみならず、親関数（さらに親関数を呼び出している関数など、再帰的にすべての関数）も同様です。  
そのため、「親関数 2」ログは出力されません。


#### 副作用

この方法はマクロ全体の実行を終了するため、  
いくつかの副作用があります。

①VBAが占有していたメモリが解放される。

変数の値がクリアされ、すべてのオブジェクトが初期化されます。  

そのため、モジュール変数に値が格納されていることが想定されたマクロの実行などは、  
`End`実行の後は正常に動作しません（再度変数に値をセットするところから始めないといけません）。

②Openしていたファイルが閉じられる。

編集したり読み取りしたりしていたファイルが強制的に遮断されます。  
書き込み内容がブツ切れにならないように、  
ファイルI/Oに関連した処理で`End`を使用するのは（可能であれば）避けたほうが良いでしょう。

### Quit（Excelアプリを終了して閉じる）

#### ABOUT

![Quit（Excelアプリを終了して閉じる）](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1641001541/learnerBlog/vba-question-006-kinds-of-end-process/kinds-of-end-process4_g8xqiq.png)

Excelアプリ自体を閉じる方法です。

`Application.Quit`を実行することで、  
マクロだけでなくExcelアプリケーション自体を終了することになるので、  
他に開いているブックを保存するか・しないか、  
また、終了前の処理は何を行うべきか、などをあわせて  
検討すると良いでしょう。

#### コード

```vb
Sub 親関数()
    Debug.Print "親関数 " & 1
    Call 子関数
    Debug.Print "親関数 " & 2
End Sub

Sub 子関数()
    Debug.Print "子関数 " & 1
    Application.Quit
    Debug.Print "子関数 " & 2
    Debug.Print "子関数 " & 3
End Sub
```

親関数を実行すると、  
子関数の`Quit`実行行で、Excelが閉じられます。

それ以降の処理は実行されないままExcelが終了されます。

## 終わりに

「マクロを終了する」という一つの要望に対しても、  
付帯する様々な結果があることがわかります。

終了処理が不十分なままマクロを終了しないように注意して、  
適切な方法を選択するようにできれば、  
マクロの品質が上がるでしょう。
