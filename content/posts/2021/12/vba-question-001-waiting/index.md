---
title: "[教えて！VBA] 第1回 処理の途中で待機時間（Sleep）を設けるにはどうすればいいの？？"
author: dede-20191130
date: 2021-12-12T10:41:37+09:00
slug: vba-question-001-waiting
draft: false
toc: true
featured: false
tags: ["VBA"]
categories: ["プログラミング"]
archives:
    - 2021
    - 2021-12
---

## この記事について

{{< box-with-title title="かんたんな概要と結論" >}} 
    VBAと他のアプリケーションを連携したり、Webサイトから情報を取得する際に、<br>
    待機時間を設ける必要があるケースがあります。<br>
    <code>wait</code>と<code>sleep</code>の二種類の方法があり、<br>
    基本的にどちらを使用しても実現できます。
{{< /box-with-title >}}

こんにちは、dedeです。

この記事では、  
VBAマクロに関する質問のうち、  
皆が疑問に思っているトピックについて解説いたします。

今回は、
{{< colored-span color="#fb9700" >}}処理の途中で待機時間を設けるメリットとその方法について{{< /colored-span >}}  
を解説いたします。

レベル：<b>中級者向け</b>


## 待機はなぜ必要？

そもそもなぜ待機時間を設ける必要があるのか？

### ①他のアプリケーションの処理が終わるのを待つため

VBAから他のアプリケーションの処理をトリガーする場合、  
アプリケーションの種類によっては同期的に処理できない  
（トリガーした方のアプリが終了してから次の処理を行うことができない）場合があります。

```vb

Call runSomethingApp()'他のアプリを起動して処理を行う

Set result = getSomethingAppResult()'※※ まだ他のアプリケーションの処理が終わっていないので、想定した結果が得られない

```

そのため、次のようにして  
アプリケーションの処理が十分に終えることができるような待機時間を儲けます。

```vb

Call runSomethingApp()'他のアプリを起動して処理を行う

Call setTaikiJikan()'待機時間呼び出し関数

Set result = getSomethingAppResult()'待機し、他のアプリが処理完了したので、正常に処理ができる

```

あるいは、次のようにして処理が終わるまでループを回すほうが安全でしょう。

```vb

Call runSomethingApp()'他のアプリを起動して処理を行う

Do
    Set result = getSomethingAppResult()'アプリ結果取得
    If Not result Is Nothing Then Exit Do'結果が得られれば次の処理へ
    If counter > 10 Then
        MsgBox "アプリケーションは正常に動作しませんでした"
        Exit Function
    End If
    Call setTaikiJikanOneSecond()'一秒待機する
    counter = counter + 1
Loop

```

VBAから他のアプリケーションを発火させる際は、  
同期的に処理できるか、非同期的な処理となるかという点は  
注意するべきでしょう。

### ②Webページのスクレイピングで、ページが表示されるまで待機するため

スクレイピング目的で  
プログラムからWebにアクセスした場合、  
通常のブラウジングと同じようにページがすべて表示されるまで待機する必要があります。

(下記、IEに対する操作自動化のコード例ですが、他のツールを使用したChromeなどへのスクレイピングでも同様です)


```vb
Dim objIE As InternetExplorer

Set objIE = New InternetExplorer
objIE.Visible = True

'指定したURLのページを表示する
objIE.navigate "https://dede-20191130.github.io/learnerBlog/about/"

'完全にページが表示されるまで待機する（想定したHTMLを得られるようにするため）
Do While objIE.Busy = True Or objIE.readyState <> 4
    Call setTaikiJikanOneSecond()'一秒待機する
    DoEvents
Loop

doSomething()'次の処理

```



## 待機の方法

### Excel VBAとAccess VBAでは別々の方法

待機を実装するための方法が、  
ExcelとAccessとで微妙に異なるので、  
それぞれ記載したいと思います。

### Excel VBA

① Application.Waitを使用する

```vb
Application.Wait(Time)
```

ApplicationはExcelアプリのインスタンス、  
つまり、Excelアプリがもともと持っているWaitという関数を使用するかたちです。

`Time`には、Excelで使用できる時間形式が入ります（00:00:10など）。  
そのため、ミリ秒単位での待機をしようとすると工夫が必要です。  
[➩ https://www.higashisalary.com/entry/vba-wait-ms](https://www.higashisalary.com/entry/vba-wait-ms)

機能は、指定した`Time`まで処理を待機するというものです。  

つまり、`Now`関数と組み合わせることで、  
「〇〇秒経過するまで待機」という処理を実現することができます。  

```
doSomething...

Call Application.Wait（Now + TimeValue（"0:00:10"））'10秒だけ待機します。

doSomethingAfter10...
```

② Windows APIのSleep関数を使用する

```vb
'モジュールの先頭
#If VBA7 Then
 Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr) '64ビットOfficeの場合
#Else
 Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds as Long) '32ビットOfficeの場合
#End If

Sub SleepTest()
    doSomething...
    Sleep 10000 '10秒だけ待機します。
    doSomethingAfter10...
End Sub
```

Windows API(Application Programming Interface)を使用します。  
Windows APIは、Windowsでよく使用される機能がまとまったライブラリです。

上記コード例のように、  
モジュールのはじめで`Public Declare......`というようにAPIを読み込む宣言をします。

`Sleep`関数の引数には、ミリ秒の指定が入ります。



### Access VBA

① Windows APIのSleep関数を使用する

上記でご説明したように、  
`Sleep`はWindows APIであり、Officeの各アプリケーションの固有の機能ではありません。

そのため、  
ExcelでもAccessでも（あるいはOutLookなどでも）、読み込みさえしておけば自由に使用することができます。

### WaitとSleep、どちらを使う？

Excel以外のOfficeのVBA環境では`Sleep`を使用すれば問題ないかと思います。

`Sleep`に対する`Wait`の優位性としては、  
時間の正確さがあります。

`Sleep`はマシンのプロセッサの時間刻みに依存しており、  
マシンごとにわずかに異なる可能性のある時間遅延を計算する恐れがあります。

といっても、待機として用いる際に  
そこまで正確な時間計算が求められるかと言われれば怪しいため、  
ミリ秒の待機を簡潔に書きたければ`Sleep`、  
外部APIの読み込みの記述が面倒くさければ`Wait`で問題ないでしょう。

## デモ（Webサイトのスクレイピング）

```vb
Sub デモ_本ブログの自己紹介欄から情報取得()
    Dim objIE As InternetExplorer
    
    Set objIE = New InternetExplorer
    objIE.Visible = False
    objIE.navigate "https://dede-20191130.github.io/learnerBlog/about/" '自己紹介ページを開く

    '完全にページが表示されるまで待機する
    Do While objIE.Busy = True Or objIE.readyState <> 4
        Application.Wait (Now + TimeValue("0:00:01")) '//一秒待機
        DoEvents
    Loop
    
    With objIE.Document.getElementsByTagName("table")(0)
        Debug.Print .Rows(3).Cells(1).innerText '//githubのURLを取得
        Debug.Print .Rows(4).Cells(1).innerText '//twitterのURLを取得
    End With
    
    objIE.Quit 'クローズ
    Set objIE = Nothing
    
End Sub

```

上記コードを実行すると、  
次のようにイミディエイトウィンドウにURLが表示されます。

```vb
https://github.com/dede-20191130 
https://twitter.com/D20191130 
```




