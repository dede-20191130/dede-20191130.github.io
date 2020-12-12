---
title: "[VBA] Excelの複数シートをループを使わず一行の処理で非表示・再表示にすることはできる？？"
author: dede-20191130
date: 2020-12-12T11:49:05+09:00
slug: Sheet-Array-Visible
draft: false
toc: true
featured: false
tags: []
categories: []
archives:
    - 2020
    - 2020-12
---

## この記事について

Excelのシートに対する一括処理は、  
次のようにシート名の配列としてSheetsオブジェクトを生成することで  
ループを使わず一行の処理で実装することが可能。

```vb
'社員シートと勤怠シートをまとめて削除する
Call Sheets(Array("社員", "勤怠")).Delete
```

このようにして、  
複数のシートの非表示・再表示（Visibleプロパティの変更）を、  
ループを使わず一行の処理でできるかどうか？という命題がある。

探したけれど、これについて言及している日本語サイトは存在しなかったので、  
検証した結果を記したい。  

ちなみにstackoverflowなどには  
これについていくつか記事が存在した。  
[https://stackoverflow.com/questions/55776641/excel-vba-unhiding-sheets-in-array-run-time-error-13-type-mismatch](https://stackoverflow.com/questions/55776641/excel-vba-unhiding-sheets-in-array-run-time-error-13-type-mismatch)  
[https://answers.microsoft.com/en-us/msoffice/forum/msoffice_excel-msoffice_custom-mso_2007/hide-and-unhide-and-array-of-worksheets/d9e14ebe-eb5f-4339-bff8-8354afe79b64?auth=1](https://answers.microsoft.com/en-us/msoffice/forum/msoffice_excel-msoffice_custom-mso_2007/hide-and-unhide-and-array-of-worksheets/d9e14ebe-eb5f-4339-bff8-8354afe79b64?auth=1)

## 検証

### 非表示にすることはできる？

問題なくできる。

```vb
Sub hideSheetArray()
    ActiveWorkbook.Sheets(Array("accounting", "finance")).Visible = False
End Sub
```

上の関数は、  
現在開いているブックのaccountingシートとfinanceシートを非表示にする。

### 再表示することはできる？

こちらは {{< colored-span color="#fb9700" >}}できない{{< /colored-span >}}。

```vb
Sub unHideSheetArray_notWorking()
    ActiveWorkbook.Sheets(Array("accounting", "finance")).Visible = True
End Sub
```

UI上でも再表示は一つずつしかできないので、  
Visibleプロパティの特徴と考えるしかなさそう。

![再表示のUI](./image01.png)

### 一括で再表示にするにはループを使うしかない

残念ながらそのようである。

```vb
Sub unHideSheetArray()
    Dim sh As Worksheet
    
    For Each sh In ActiveWorkbook.Sheets(Array("accounting", "finance"))
        sh.Visible = True
    Next sh
    
End Sub
```

