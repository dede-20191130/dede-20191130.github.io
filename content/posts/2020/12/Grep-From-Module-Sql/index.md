---
title: "[VBA, PowerShell] Accessのモジュール・クラスやクエリのSQLから特定文字列を抽出するためのテクニック"
author: dede-20191130
date: 2020-12-19T03:18:51+09:00
slug: Grep-From-Module-Sql
draft: false
toc: true
featured: false
tags: ['Access', 'VBA','PowerShell','自作']
categories: ['プログラミング']
archives:
    - 2020
    - 2020-12
---


## この記事について

AccessのVBAツールを作成していると、  
主にリファクタリングや機能追加のタイミングで、  
モジュール、クラス、あるいはクエリのソース（SQL）から  
特定の文字列をサーチして抽出したいことがあるかもしれない。

モジュール・クラスの場合はVBエディタ画面でctrl+Fを実行すれば可能だが、一覧でヒットした箇所を表示できないため、  
全体を把握するのが大変だ。  
また、クエリのSQLからサーチする機能は無いかと思う。

そのため、私は以下のようにして、  
一度それぞれのデータをファイルとして取り出したあとに  
PowerShellのコマンドでLinuxのGrepコマンド風に文字列を抽出している。

## 作成環境

- Windows10 Home
- MSOffice 2019
- PowerShell 5.1

## テクニック

### それぞれのデータをファイルとして取り出す

#### about

モジュールやクラスは、  
VBAを用いてすべて一括でファイルとして取り出すことができる。  
VBComponentオブジェクトのExportメソッドを使用すれば良い。



{{< inner-article-div color="#fb9700" >}}ただ、RubberDuckなどの拡張アドインをいれている場合は、  <br>
そちらのエクスポート機能を使用したほうが手早い。{{< /inner-article-div >}}

また、  
この関数で、    
クエリのSQLも同時にsqlファイルとして取り出す。

#### コード

```vb
'******************************************************************************************
'*関数名    ：exportCodesSQLs
'*機能      ：モジュール・クラスのコード及びクエリのSQLの出力
'*引数      ：
'******************************************************************************************
Sub exportCodesSQLs()
    
    '定数
    Const FUNC_NAME As String = "XXX"
    
    '変数
    Dim outputDir As String
    Dim vbcmp As Object
    Dim fileName As String
    Dim ext As String
    Dim qry As QueryDef
    Dim qName As String
    
    
    
    On Error GoTo ErrorHandler
    
    outputDir = _
        Access.CurrentProject.Path & _
        "\" & _
        "src_" & _
        Left(Access.CurrentProject.Name, InStrRev(Access.CurrentProject.Name, ".") - 1)
    If Dir(outputDir) = "" Then MkDir outputDir
    
    'モジュール・クラスの出力
    For Each vbcmp In VBE.ActiveVBProject.VBComponents
        With vbcmp
            '拡張子
            Select Case .Type
            Case 1
                ext = ".bas"
            Case 2, 100
                ext = ".cls"
            Case 3
                ext = ".frm"
            End Select
                        
            fileName = .Name & ext
            fileName = gainStrNameSafe(fileName) 'ファイル名に使用できない文字を置換
            If fileName = "" Then GoTo ExitHandler
            
            'output
            .Export outputDir & "\" & fileName
            
        End With
    Next vbcmp
    
    'SQLの出力
    With CreateObject("Scripting.FileSystemObject")
        For Each qry In CurrentDb.QueryDefs
            Do
                qName = gainStrNameSafe(qry.Name) 'ファイル名に使用できない文字を置換
                If qName = "" Then GoTo ExitHandler
                
                If qName Like "Msys*" Then Exit Do 'システム関連クエリは除外
                
                With .CreateTextFile(outputDir & "\" & qName & ".sql")
                    .write qry.SQL
                    .Close
                End With
            Loop While False
        Next qry
    End With

ExitHandler:

    Exit Sub
    
ErrorHandler:

    MsgBox "エラーが発生したため、マクロを終了します。" & _
           vbLf & _
           "関数名：" & FUNC_NAME & _
           vbLf & _
           "エラー番号：" & err.Number & vbNewLine & _
           err.description, vbCritical, "マクロ"
        
    GoTo ExitHandler
        
End Sub




'******************************************************************************************
'*関数名    ：gainStrNameSafe
'*機能      ：ファイル名に使用できない文字をアンダースコアに置換する
'*引数      ：対象の文字列
'*戻り値    ：置換後文字列
'******************************************************************************************
Public Function gainStrNameSafe(ByVal s As String) As String
    
    '定数
    Const FUNC_NAME As String = "gainStrNameSafe"
    
    '変数
    Dim x As Variant
    
    On Error GoTo ErrorHandler

    gainStrNameSafe = ""
    
    For Each x In Split("\,/,:,*,?,"",<,>,|", ",") 'ファイル名に使用できない文字の配列
        s = Replace(s, x, "_")
    Next x
    
    gainStrNameSafe = s

ExitHandler:

    Exit Function
    
ErrorHandler:

    MsgBox "エラーが発生したため、マクロを終了します。" & _
           vbLf & _
           "関数名：" & FUNC_NAME & _
           vbLf & _
           "エラー番号：" & err.Number & vbNewLine & _
           err.description, vbCritical, "マクロ"
        
    GoTo ExitHandler
        
End Function
```

exportCodesSQLsを実行すれば、  
Accessファイルの置かれているディレクトリの「src_Accessファイル名」フォルダに  
各種ソースファイルが格納される。

![エクスポートされたファイル](./image01.png)

### PowerShellのコマンドで文字列を抽出する

#### about

PowerShellを起動し、  
エクスポートされたファイルのディレクトリに移動する。  

LinuxのGrep風に  
文字列をサーチして一覧表示するコマンドは、  
以下のようにした。

```PowerShell
Get-ChildItem | ForEach-Object{ Write-Output  ($_.Name + "`r`n------") ; (Get-Content $_   | Select-String "ここにサーチしたい文字列を記入する"  )  | ForEach-Object{Write-Output ($_.lineNumber.Tostring() + ":" + $_) } ;Write-Output "------"  } 
```

最初にサーチ対象ファイル名が表示され、ヒットした行番号とその行の文字列が出力される。  
これをファイルごとにループさせる。

#### 例

例えば、全てのファイルから  
「ID」という文字列をサーチして一覧で表示したい場合、  
次のようにコマンドを実行する。

```PowerShell
Get-ChildItem | ForEach-Object{ Write-Output  ($_.Name + "`r`n------") ; (Get-Content $_   | Select-String "ID"  )  | ForEach-Object{Write-Output ($_.lineNumber.Tostring() + ":" + $_) } ;Write-Output "------"  }  
```

結果は例えばこのようになる。  

![Grep結果01](./image02.png)

![Grep結果02](./image03.png)



