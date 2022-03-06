---
title: "EXCEL VBA: YOU MAY HAVE A ERROR ON CHANGING PRINTAREA DYNAMICALLY & HOW TO AVOID IT"
author: dede-20191130
date: 2021-01-15T23:38:43+09:00
slug: Change-PageSetup-PrintAria
draft: false
toc: true
featured: false
tags: ['VBA','Excel']
categories: ['programming','Trouble Shooting']
vba_taxo: specification
archives:
    - 2021
    - 2021-01
---

{{< help-me-improve-lang  >}}

## ABOUT THIS ARTICLE

You may want your implementaion to change print area of a worksheet dynamically with `PageSetup.PrintArea` according to some conditions.

At that time, you may be subjected to a unexpected error if you are not careful with {{< colored-span color="#fb9700" >}}the Cell Reference Style{{< /colored-span >}} of Excel App.

In this article, I'd like to describe the case the error occurs and two methods how to avoid it.

You can download Excel file created for explanation, and view its source code from [<span id="srcURL"><u>here</u></span>](https://github.com/dede-20191130/My_VBA_Tools/tree/master/Public/2021/01/Change-PageSetup-PrintAria/en)!


## VERIFICATION ENVIRONMENT

Windows 10 Home(64bit)  
MSOffice 2016

## THE CASE

### ABOUT

Let's assume that the worksheet of your Excel book has its print area set.

The program performs the processing to extend the area to one more line below.  
e.g. If initial area is *$A$1:$E$5*, after executing it'll chnage to *$A$1:$E$6*.


### CODE

```vb {hl_lines=[21]}
'******************************************************************************************
'*Function :it's a function Before Modified
'*          extend PrintArea to one line below
'******************************************************************************************
Public Sub changePrintAreaBeforeModified()
    
    'Consts
    Const FUNC_NAME As String = "changePrintAreaBeforeModified"
    
    'Vars
    Dim prePrintAreaAddress As String
    
    On Error GoTo ErrorHandler
    
    With ThisWorkbook.Worksheets(1)
    
        'Current Print Area Address
        prePrintAreaAddress = .PageSetup.PrintArea
        
        'extend PrintArea to one line below
        .PageSetup.PrintArea = .Range(prePrintAreaAddress).Resize(.Range(prePrintAreaAddress).Rows.Count + 1).Address
        
        Debug.Print .PageSetup.PrintArea
        
    End With

ExitHandler:

    Exit Sub
    
ErrorHandler:

    MsgBox "An error has occurred and the macro will be terminated." & _
           vbLf & _
           "Func Name:" & FUNC_NAME & _
           vbLf & _
           "Error No." & Err.Number & vbNewLine & _
           Err.Description, vbCritical, SOURCE_NAME
        
    GoTo ExitHandler
        
End Sub
```

## A ERROR OCCURED IF R1C1 USED

### DETAIL

Above code works if you use A1 reference style.  
But if you use R1C1 one, you'll get a error on the highlighted line.


![The Error](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1646442992/learnerBlog/Change-PageSetup-PrintAria/en/%E3%82%B9%E3%82%AF%E3%83%AA%E3%83%BC%E3%83%B3%E3%82%B7%E3%83%A7%E3%83%83%E3%83%88_2022-03-05_101428_q8fkbq.png)

### CAUSE

`PrintArea` property of `PageSetup` object gets different addresses depending on application's reference style at the time the function is executed.

- A1 style address string if A1 used
- R1C1 style address string if R1C1 used

And the address string that `Worksheet.Range` requires must be A1 style, not allowed if R1C1 style.

Thus, `prePrintAreaAddress` stores the R1C1 styled address and the error occurs when `Worksheet.Range` gets `prePrintAreaAddress` as a argument.



## HOW TO AVOID

### PATTERN 1. SWITCH APPLICATION'S REF STYLE ITSELF

```vb {hl_lines=["19-20"]}
'******************************************************************************************
'*Function :it's a function after midification of pattern No.1
'*          extend PrintArea to one line below
'******************************************************************************************
Public Sub changePrintAreaModified01()
    
    'Consts
    Const FUNC_NAME As String = "changePrintAreaModified01"
    
    'Vars
    Dim prePrintAreaAddress As String
    Dim currentStyle As XlReferenceStyle

    On Error GoTo ErrorHandler
    
    With ThisWorkbook.Worksheets(1)
        
        'change the reference style to A1 style
        currentStyle = Application.ReferenceStyle
        Application.ReferenceStyle = xlA1
        
        'Current Print Area Address
        prePrintAreaAddress = .PageSetup.PrintArea
        
        'extend PrintArea to one line below
        .PageSetup.PrintArea = .Range(prePrintAreaAddress).Resize(.Range(prePrintAreaAddress).Rows.Count + 1).Address
        
        Debug.Print .PageSetup.PrintArea
        
        'restore the reference style
        Application.ReferenceStyle = currentStyle
        
    End With

ExitHandler:

    Exit Sub
    
ErrorHandler:

    MsgBox "An error has occurred and the macro will be terminated." & _
           vbLf & _
           "Func Name:" & FUNC_NAME & _
           vbLf & _
           "Error No." & Err.Number & vbNewLine & _
           Err.Description, vbCritical, SOURCE_NAME
        
    GoTo ExitHandler
        
End Sub

```

Application's reference style is switched to A1 forcibly before setting `PrintArea` and restored after it.

The disadvantage is that,  
when there is a time-consuming process between switching or when the function is called many times,  
the user may see a screen flicker during the process because of switching the ref-style.






### PATTERN 2. MODIFY THE REF STYLE OF THE OBJECT'S ADDRESS TO A1 STYLE

With `Application.ConvertFormula`, the address string stored a variable can be changed to A1 ref-style without involving Application.ReferenceStyle.


```vb {hl_lines=["22"]}
'******************************************************************************************
'*Function :it's a function after midification of pattern No.2
'*          extend PrintArea to one line below
'******************************************************************************************
Public Sub changePrintAreaModified02()
    
    'Consts
    Const FUNC_NAME As String = "changePrintAreaModified02"
    
    'Vars
    Dim prePrintAreaAddress As String
    
    On Error GoTo ErrorHandler
    
    With ThisWorkbook.Worksheets(1)
    
        'Current Print Area Address
        prePrintAreaAddress = .PageSetup.PrintArea
        
        'modify the address of prePrintAreaAddress to xlR1C1 style
        '** it doesn't change application's reference style
        If Application.ReferenceStyle = xlR1C1 Then prePrintAreaAddress = Application.ConvertFormula(prePrintAreaAddress, xlR1C1, xlA1)
        
        'extend PrintArea to one line below
        .PageSetup.PrintArea = .Range(prePrintAreaAddress).Resize(.Range(prePrintAreaAddress).Rows.Count + 1).Address
        
        Debug.Print .PageSetup.PrintArea
        
    End With

ExitHandler:

    Exit Sub
    
ErrorHandler:

    MsgBox "An error has occurred and the macro will be terminated." & _
           vbLf & _
           "Func Name:" & FUNC_NAME & _
           vbLf & _
           "Error No." & Err.Number & vbNewLine & _
           Err.Description, vbCritical, SOURCE_NAME
        
    GoTo ExitHandler
        
End Sub

```

## AT THE END

THe latter resolution is more flexible and user-friendly, I think.

