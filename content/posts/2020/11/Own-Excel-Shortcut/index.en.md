---
title: "EXCEL VBA: HOMEMADE EXCEL SHORTCUTS TO IMPROVE WORK EFFICIENCY"
author: dede-20191130
date: 2020-11-05T00:48:58+09:00
slug: Own-Excel-Shortcut
draft: false
toc: true
tags: ['Excel', 'VBA','homemade']
categories: ['Application', 'programming']
vba_taxo: help_develop
archives:
    - 2020
    - 2020-11
---

{{< help-me-improve-lang  >}}

## ABOUT THIS ARTICLE

I think there are quite a lot of cases to post text from Excel to another file or document, cut and paste shapes such as rectangle and arrow on Excel sheet, and create a very simple flowchart or schematic diagram.

In these cases, what improved the speed of work was my homemade shortcuts which are not not originally included in Excel, so I'm going to introduce them.



## LIST

List below are the shortcuts to be introduced



|Functionality|When To Use|Assigned Key|
|--|--|--|
|copy only text from target cell|when gettting sentence in the cell <br/>without double quotes at both ends|Ctrl + Shift + K|
|move the selected object to front or back|when creating a little complicated diagram etc.|Ctrl + Shift + B|
|pause Excel Events|when opening Excel books with macro, <br/>without running event processing <br/>which automatically open form|Ctrl + Shift + M|


## HOW TO REGISTER MACROS FOR USING HOMEMADE SHORTCUT

1. open VBE.
2. write some procedures in standard module.
3. back to sheet, and open Macro setting Screen by pressing Alt F8.
4. select target procedure and register shortcut key to call it via options.




## WHAT KEY TO REGISTER?

The most convenient way to register shortcuts is to use a key that has not yet been reserved or is less used normally.  
I suggest Ctrl + Shift + K, M, N.



## SEPARATE DESCRIPTION

### SHORTCUT TO COPY ONLY TEXT FROM TARGET CELL

When we try to copy cell in sheet and paste it to another application such as notepad, it goes along with some extra stuff, i.e. Line Feed and double quotes.  
They are sometimes botherring our task.

Excel App has a function of pasting only values but doesn't have copying only values, I guess that's what causes our small troubles.

Below is a shortcut to improve this.


#### CODE

```vb
'******************************************************************************************
'*Function :copy activecell's content to clipboard
'******************************************************************************************
Public Sub copyCellValueToCB()

    'Const
    Const FUNC_NAME As String = "copyCellValueToCB"

    'Vars

    On Error GoTo ErrorHandler

    'store text to clipboard
    With CreateObject("Forms.TextBox.1")
        .MultiLine = True
        .Text = CStr(ActiveCell.Value)
        .SelStart = 0
        .SelLength = .TextLength
        .Copy
    End With

ExitHandler:

    Exit Sub

ErrorHandler:

        MsgBox "An error has occurred and the macro will be terminated." & _
           vbLf & _
           "Func Name:" & FUNC_NAME & _
           vbLf & _
           "Error No." & Err.Number & vbNewLine

        GoTo ExitHandler

End Sub
```


### SHORTCUT TO MOVE THE SELECTED OBJECT TO THE FRONT OR BACK

When creating a simple flowchart, schematic diagram, and organization chart,  
probably there are a case that we want to move shapes to the front or back against the other shapes in order to adjust the overlap of them.

In that case, calling the processing of 'bring to front' by Right-click is slow, and this shortcut reduce time of it.



#### CODE

```vb
'******************************************************************************************
'******************************************************************************************
Public Sub ZOrderToFront()
    
    'Const
    Const FUNC_NAME As String = "ZOrderToFront"
    
    'Vars
    
    On Error GoTo ErrorHandler
    
    Selection.ShapeRange.ZOrder msoBringToFront

ExitHandler:

    Exit Sub
    
ErrorHandler:
    
    If Err.Number = 438 Then
        MsgBox "Plrease run after selecting target object.", vbExclamation, "Warning"
    Else
        MsgBox "An error has occurred and the macro will be terminated." & _
           vbLf & _
           "Func Name:" & FUNC_NAME & _
           vbLf & _
           "Error No." & Err.Number & vbNewLine

    End If
    GoTo ExitHandler
        
End Sub

```

If you try to make `ZOrderToBack`, replace:  
```vb
Selection.ShapeRange.ZOrder msoSendToBack
```




### SHORTCUT TO PAUSE EXCEL EVENTS 

Useful in cases below:  
- When editting Excel book with macro, you want to launch it without opening event procedure.
- When switching active worksheet, some event may be executed and you fell it troublesome.

#### CODE

The entire process divides into two parts: caller funtion and core process in a form.

##### CALLER: IN STANDARD MODULE

```vb

'in Tools.bas

'******************************************************************************************
'*Function      :Disable All Excel Events during displaying a F_invalidateEvents form
'******************************************************************************************
Public Sub invalidateEvents()

    'Const
    Const FUNC_NAME As String = "invalidateEvents"

    On Error GoTo ErrorHandler

    'open the form
    F_invalidateEvents.Show vbModeless

ExitHandler:

    Exit Sub

ErrorHandler:

        MsgBox "An error has occurred and the macro will be terminated." & _
           vbLf & _
           "Func Name:" & FUNC_NAME & _
           vbLf & _
           "Error No." & Err.Number & vbNewLine

        GoTo ExitHandler

End Sub

```

##### FORM: CORE PROCESS

```vb

' in F_invalidateEvents.frm

Option Explicit


'******************************************************************************************
'******************************************************************************************
Private Sub UserForm_Initialize()

    'Const
    Const FUNC_NAME As String = "UserForm_QueryClose"

    'Vars

    On Error GoTo ErrorHandler

    'disable events
    Application.EnableEvents = False

ExitHandler:

    Exit Sub

ErrorHandler:

        MsgBox "An error has occurred and the macro will be terminated." & _
           vbLf & _
           "Func Name:" & FUNC_NAME & _
           vbLf & _
           "Error No." & Err.Number & vbNewLine

        GoTo ExitHandler

End Sub


'******************************************************************************************
'******************************************************************************************
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    'Const
    Const FUNC_NAME As String = "UserForm_QueryClose"

    'Vars

    On Error GoTo ErrorHandler

    'enable events
    Application.EnableEvents = True

ExitHandler:

    Exit Sub

ErrorHandler:

        MsgBox "An error has occurred and the macro will be terminated." & _
           vbLf & _
           "Func Name:" & FUNC_NAME & _
           vbLf & _
           "Error No." & Err.Number & vbNewLine

        GoTo ExitHandler

End Sub

'******************************************************************************************
'******************************************************************************************
Private Sub CommandButton_Close_Click()

    'Const
    Const FUNC_NAME As String = "CommandButton_Close_Click"

    'Vars

    On Error GoTo ErrorHandler

    'close form
    Unload F_invalidateEvents

ExitHandler:

    Exit Sub

ErrorHandler:

        MsgBox "An error has occurred and the macro will be terminated." & _
           vbLf & _
           "Func Name:" & FUNC_NAME & _
           vbLf & _
           "Error No." & Err.Number & vbNewLine

        GoTo ExitHandler

End Sub

```

#### DEMO

After calling the form, it disable all events initially, and you can restore it by pressing close button.

![F_invalidateEvents Form](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1645934631/learnerBlog/Own-Excel-Shortcut/en/%E3%82%B9%E3%82%AF%E3%83%AA%E3%83%BC%E3%83%B3%E3%82%B7%E3%83%A7%E3%83%83%E3%83%88_2022-02-27_125835_qr9ubk.png)


## AT THE END

I'm going to update this article if I create a new handy shortcut.
