---
title: "VBA & POWERSHELL: THE TECHNIQUE TO EXTRACT A SPECIFIC STRING FROM MODULES, CLASSES OR SQL OF QUERIES IN ACCESS DATABASE"
author: dede-20191130
date: 2020-12-19T03:18:51+09:00
slug: Grep-From-Module-Sql
draft: false
toc: true
featured: false
tags: ['Access', 'VBA','PowerShell','HOMEMADE']
categories: ['programming']
vba_taxo: technique_develop
archives:
    - 2020
    - 2020-12
---

{{< help-me-improve-lang  >}}

## ABOUT THIS ARTICLE

When creating a Access VBA tool, you may search and extract a specific string from modules, classes, or query source SQL,   
mainly in order to refactor code or add a new feature.

For modules and classed, on VB Editor screen you can search it by pressing ctrl + F, but you can't get a list of matches and it's difficult to grasp the total result.  
And, there is not a function to search a string from query source SQL in Access as of 2022.

Therefore, under Windows, I export each data as a file and then extract target string with powershell command like Linux's Grep command, as explained below.

## CREATION ENVIRONMENT

- Windows10 Home
- MSOffice 2019
- PowerShell 5.1

## TECHNIQUE

### EXPORT EACH DATA AS A FILE

#### ABOUT

You can export all of modules and classes by running VBA function at once. To do so, you use `Export` method of the VBComponent object.



{{< inner-article-div color="#fb9700" >}}
However, if you introduced Add-ins such as RubberDuck,   
it may be easy to use its own export function.
{{< /inner-article-div >}}

And, through this function, you'll find query sources are exported as a sql file simultaneously.


#### CODE

```vb
'******************************************************************************************
'*Function      :output codes of module and class, and query source SQLs
'******************************************************************************************
Sub exportCodesSQLs()
    
    'Consts
    Const FUNC_NAME As String = "exportCodesSQLs"
    
    'Vars
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
    If Dir(outputDir, vbDirectory) = "" Then MkDir outputDir
    
    'output modules, classes
    For Each vbcmp In VBE.ActiveVBProject.VBComponents
        With vbcmp
            'set extension
            Select Case .Type
            Case 1
                ext = ".bas"
            Case 2, 100
                ext = ".cls"
            Case 3
                ext = ".frm"
            End Select
                        
            fileName = .Name & ext
            fileName = gainStrNameSafe(fileName) 'replace some charactors which aren't allowed to use for a file name.
            If fileName = "" Then GoTo ExitHandler
            
            'output
            .Export outputDir & "\" & fileName
            
        End With
    Next vbcmp
    
    'output query sources
    With CreateObject("Scripting.FileSystemObject")
        For Each qry In CurrentDb.QueryDefs
            Do
                qName = gainStrNameSafe(qry.Name) 'replace some charactors which aren't allowed to use for a file name
                If qName = "" Then GoTo ExitHandler
                
                If qName Like "Msys*" Then Exit Do 'exclude queries related MS system
                
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

    MsgBox "An error has occurred and the macro will be terminated." & _
           vbLf & _
           "Func Name:" & FUNC_NAME & _
           vbLf & _
           "Error No." & Err.Number & vbNewLine & _
           Err.Description, vbCritical, "Macro"
        
    GoTo ExitHandler
        
End Sub




'******************************************************************************************
'*Function      :replace some charactors with a underscore and return replaced string. The charactors aren't allowed to use for a file name.
'*Arg           :target string
'*Return        :replaced  string
'******************************************************************************************
Public Function gainStrNameSafe(ByVal s As String) As String
    
    'Consts
    Const FUNC_NAME As String = "gainStrNameSafe"
    
    'Vars
    Dim x As Variant
    
    On Error GoTo ErrorHandler

    gainStrNameSafe = ""
    
    For Each x In Split("\,/,:,*,?,"",<,>,|", ",") 'array of chars not to be used
        s = Replace(s, x, "_")
    Next x
    
    gainStrNameSafe = s

ExitHandler:

    Exit Function
    
ErrorHandler:

    MsgBox "An error has occurred and the macro will be terminated." & _
           vbLf & _
           "Func Name:" & FUNC_NAME & _
           vbLf & _
           "Error No." & Err.Number & vbNewLine & _
           Err.Description, vbCritical, "Macro"
        
    GoTo ExitHandler
        
End Function
```

After running `exportCodesSQLs`,   
you'll find all source files have been stored into 'src_ + Access file name' folder  
directly under the folder where the Access file is located.



![Exported Files](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1645968367/learnerBlog/Grep-From-Module-Sql/en/%E3%82%B9%E3%82%AF%E3%83%AA%E3%83%BC%E3%83%B3%E3%82%B7%E3%83%A7%E3%83%83%E3%83%88_2022-02-27_222415_lkhikd.png)

### EXTRACT TARGET STRING BY POWERSHELL COMMAND

#### ABOUT

Launch PowerShell, and move the folder in which exported files are.

The following is a command to search a specific string like Linux's Grep command and show the list. 



```PowerShell
Get-ChildItem | ForEach-Object{ Write-Output  ($_.Name + "`r`n------") ; (Get-Content $_   | Select-String "here you write a string to be searched"  )  | ForEach-Object{Write-Output ($_.lineNumber.Tostring() + ":" + $_) } ;Write-Output "------"  } 
```

First the file name to be searched is displayed, then row number and the row's text which came up is displayed.  
The command loops this for each file.



#### DEMO

Suppose you want to search 'ITEM_CODE' from all files and look at the result list,   
execute the command as follows.



```PowerShell
Get-ChildItem | ForEach-Object{ Write-Output  ($_.Name + "`r`n------") ; (Get-Content $_   | Select-String "ITEM_CODE"  )  | ForEach-Object{Write-Output ($_.lineNumber.Tostring() + ":" + $_) } ;Write-Output "------"  }  
```

The result is:    

![Grep Result](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1645968367/learnerBlog/Grep-From-Module-Sql/en/%E3%82%B9%E3%82%AF%E3%83%AA%E3%83%BC%E3%83%B3%E3%82%B7%E3%83%A7%E3%83%83%E3%83%88_2022-02-27_222459_e28fnt.png)
