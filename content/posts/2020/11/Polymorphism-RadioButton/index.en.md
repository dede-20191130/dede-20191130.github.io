---
title: "EXCEL VBA: WITH POLYMORPHISM, BRANCHING A PROCESS WITHOUT USING IF STATEMENT"
author: dede-20191130
date: 2020-11-01T16:44:24+09:00
slug: Polymorphism-RadioButton
draft: false
toc: true
tags: ['Excel', 'VBA','HOMEMADE', 'object-oriented']
categories: ['programming']
vba_taxo: oop_others
archives:
    - 2020
    - 2020-11
---

{{< help-me-improve-lang  >}}

## ABOUT THIS ARTICLE

Hi, I'm Dede.

Here I introduce a sample of coding based on a thinking of Polymorphism in VBA.

In detail, in the sample I branched a process without using If statement, by using polymorphism implemented by `CallByName` and Tag Property of Radio Buttons in a Form.

You can download Excel file created for explanation and view its source code from [<span id="srcURL"><u>here</u></span>](https://github.com/dede-20191130/My_VBA_Tools/tree/master/Public/2020/11/Polymorphism-RadioButton/en)!



## CREATION ENVIRONMENT
Windows10  
MSOffice 2016

## PREMISE

There is a following screen and you perform a process that depends on the type of radio button by selecting one of them and pressing 'Run'.


|Name|Image|
|--|--|
|**Form**|![](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1645948671/learnerBlog/Polymorphism-RadioButton/en/%E3%82%B9%E3%82%AF%E3%83%AA%E3%83%BC%E3%83%B3%E3%82%B7%E3%83%A7%E3%83%83%E3%83%88_2022-02-27_165349_cxqijc.png)|
|**Select a radio button showing current time**|![](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1645948671/learnerBlog/Polymorphism-RadioButton/en/%E3%82%B9%E3%82%AF%E3%83%AA%E3%83%BC%E3%83%B3%E3%82%B7%E3%83%A7%E3%83%83%E3%83%88_2022-02-27_165512_edfcyz.png)|
|**Select a radio button showing user name**|![](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1645948671/learnerBlog/Polymorphism-RadioButton/en/%E3%82%B9%E3%82%AF%E3%83%AA%E3%83%BC%E3%83%B3%E3%82%B7%E3%83%A7%E3%83%83%E3%83%88_2022-02-27_165708_zstgoa.png)|
|**Select a radio button showing greeting**|![](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1645948671/learnerBlog/Polymorphism-RadioButton/en/%E3%82%B9%E3%82%AF%E3%83%AA%E3%83%BC%E3%83%B3%E3%82%B7%E3%83%A7%E3%83%83%E3%83%88_2022-02-27_165716_o8ndlx.png)|

## FORM COMPONENTS


|Name|Type|Caption|GourpName|Tag|Initial Value|
|--|--|--|--|--|--|
|rdo_showCurrent|Radio Button|show current tiem|Group01|Current|True|
|rdo_showUser|Radio Button|show user name|Group01|User|False|
|rdo_showGreeting|Radio Button|show greeting|Group01|Greeting|False|
|btn_execute|command button|Run process|--|--|--|

## OVERVIEW OF FUNCTIONS

|Name|Module/Class|Type|Functionality|
| ---- | ---- | ---- | ---- |
|btn_execute_Click| F_Main |Sub Procedure|Click event function of btn_execute<br>perform a process that depends on the type of radio button|
|btn_execute_Click_Current| clsPolymo |Function Procedure|show current time|
|btn_execute_Click_User| clsPolymo |Function Procedure|show user name|
|btn_execute_Click_Greeting| clsPolymo |Function Procedure|show greeting|

## CODE

### [btn_execute_Click]
```vb
'******************************************************************************************
'*Function :
'*Arg(1)   :
'******************************************************************************************
Private Sub btn_execute_Click()
    
    'Consts
    Const FUNC_NAME As String = "btn_execute_Click"
    
    'Vars
    Dim suffix As String
    Dim objPolymo As clsPolymo
    
    On Error GoTo ErrorHandler
    
    'instantiate a class of processings
    Set objPolymo = New clsPolymo
    
    'get selected processing flag string
    suffix = _
           WorksheetFunction.Rept(Me.rdo_showCurrent.Tag, Abs(CLng(CBool(Me.rdo_showCurrent.Value)))) & _
           WorksheetFunction.Rept(Me.rdo_showUser.Tag, Abs(CLng(CBool(Me.rdo_showUser.Value)))) & _
           WorksheetFunction.Rept(Me.rdo_showGreeting.Tag, Abs(CLng(CBool(Me.rdo_showGreeting.Value))))
    If suffix = "" Then MsgBox "Radio button selection is invalid.", vbCritical, Tool_Name: GoTo ExitHandler
    
    'call corresponding processing function
    If Not CallByName(objPolymo, FUNC_NAME & "_" & suffix, VbMethod) Then GoTo ExitHandler
    

ExitHandler:
    
    Set objPolymo = Nothing
    
    Exit Sub
    
ErrorHandler:

    MsgBox "An error has occurred and the macro will be terminated." & _
           vbLf & _
           "Func Name:" & FUNC_NAME & _
           vbLf & _
           "Error No." & Err.Number & vbNewLine & _
           Err.Description, vbCritical, Tool_Name
        
    GoTo ExitHandler
        
End Sub


```

`WorksheetFunction.Rept` function returns a string which is constructed by repeating a first parameter string for the number of times given as a second parameter.



```vb
'abcabcabc
WorksheetFunction.Rept("abc",3)
```

`Abs(CLng(CBool(Me.rdo_showCurrent.Value)))` presents **1** when target radio buttons selected, and **2** when not selected.

So, `suffix` is assigned to a string of Tag property belonging to selected radio button.

After that, a function in `clsPolymo` named `btn_execute_Click + XX` is called by using build-in `CallByName` function.  
XX indicates a tag property string.


### [btn_execute_Click_Current]

```vb
'******************************************************************************************
'*Function :show current time
'*Arg(1)   :
'*Return   :True > normal termination; False > abnormal termination
'******************************************************************************************
Public Function btn_execute_Click_Current() As Boolean
    
    'Consts
    Const FUNC_NAME As String = "btn_execute_Click_Current"
    
    'Vars
    
    On Error GoTo ErrorHandler

    btn_execute_Click_Current = False
    
    'show current time
    MsgBox "Current time: " & Now, , Tool_Name

    btn_execute_Click_Current = True
    
ExitHandler:

    Exit Function
    
ErrorHandler:

    MsgBox "An error has occurred and the macro will be terminated." & _
           vbLf & _
           "Func Name:" & FUNC_NAME & _
           vbLf & _
           "Error No." & Err.Number & vbNewLine & _
           Err.Description, vbCritical, Tool_Name
        
    GoTo ExitHandler
        
End Function
```

### [btn_execute_Click_User]

```vb
'******************************************************************************************
'*Function :show PC user name
'*Arg(1)   :
'*Return   :True > normal termination; False > abnormal termination
'******************************************************************************************
Public Function btn_execute_Click_User() As Boolean
    
    'Consts
    Const FUNC_NAME As String = "btn_execute_Click_User"
    
    'Vars
    
    On Error GoTo ErrorHandler

    btn_execute_Click_User = False
    
    With CreateObject("WScript.Network")
        'show PC user name
        MsgBox "Use name: " & .UserName, , Tool_Name
    End With

    btn_execute_Click_User = True
    
ExitHandler:

    Exit Function
    
ErrorHandler:

    MsgBox "An error has occurred and the macro will be terminated." & _
           vbLf & _
           "Func Name:" & FUNC_NAME & _
           vbLf & _
           "Error No." & Err.Number & vbNewLine & _
           Err.Description, vbCritical, Tool_Name
        
    GoTo ExitHandler
        
End Function

```

### [btn_execute_Click_Greeting]

```vb
'******************************************************************************************
'*Function :show greeting
'*Arg(1)   :
'*Return   :True > normal termination; False > abnormal termination
'******************************************************************************************
Public Function btn_execute_Click_Greeting() As Boolean
    
    'Consts
    Const FUNC_NAME As String = "btn_execute_Click_Greeting"
    
    'Vars
    
    On Error GoTo ErrorHandler

    btn_execute_Click_Greeting = False
    
    MsgBox "Hello.", , Tool_Name
    
    btn_execute_Click_Greeting = True
    
ExitHandler:

    Exit Function
    
ErrorHandler:

    MsgBox "An error has occurred and the macro will be terminated." & _
           vbLf & _
           "Func Name:" & FUNC_NAME & _
           vbLf & _
           "Error No." & Err.Number & vbNewLine & _
           Err.Description, vbCritical, Tool_Name
        
    GoTo ExitHandler
        
End Function
```

With polymorphism, which function is to be called is determined by `suffix` value, and you can branch a process without using If statement.

## A TOOL USING THE MECHANISM

Introduced by following article:  
{{< page-titled-link page="excel-a1-tool" >}}
