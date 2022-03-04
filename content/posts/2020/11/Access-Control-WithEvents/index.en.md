---
title: "Access VBA: THE THING YOU NEED TO BE AWARE OF WHEN CREATING EVENT LISTENER BY USING WITHEVENTS FOR FORM CONTROLS"
author: dede-20191130
date: 2020-11-08T10:26:29+09:00
slug: Access-Control-WithEvents
draft: false
toc: true
featured: false
tags: ['Access', 'VBA','HOMEMADE', 'object-oriented']
categories: ['Trouble Shooting', 'programming']
vba_taxo: oop_others
archives:
    - 2020
    - 2020-11
---

{{< help-me-improve-lang  >}}

## ABOUT THIS ARTICLE

I tried to create a event listener class by using `WithEvents` statement to standardize event processing for controls on a forms in MSAccess.  
At that time, when using the code which went well in MS Excel, I've gotten the result that the class side events that I've set up won't fire.

I would like to describe what happened and the two types of measures taken.

You can download the tool and view its source code from [here](https://github.com/dede-20191130/My_VBA_Tools/tree/master/Public/2020/11/Access-Control-WithEvents/en)!


## WHAT I CREATED FIRST

I made the combobox control surrounded by a circle as follows not be input manually with the keyboard.

For standardization, the processing of combobox KeyDown event is delegated to a newly created event listener class (in case the number of controls increased in the future).

For verification, I have some of the combobox events pass a textbox its event information as a log.



### SCREEN

![Screen](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1644740250/learnerBlog/Access-Control-WithEvents/en/access-control-withevents_pm0hnb.png)

### THE ROLE OF EACH OBJECT


|Name|Type|Explanation|
|--|--|--|
|cmb_withEventsTest|Combobox|Allow input only from the pullDown.<br/>Prohibit manual input|
|txt_EventLog|TextBox|display event log of the combobox|



### CODE

The code below <span style="color: red; ">does not work</span> as `myComboBox_KeyDown` function is not fired.


#### MAIN FOMR MODULE

```vb
Option Compare Database
Option Explicit

'**************************
'*MainForm
'**************************

'Const


'Variable
Private objCmbListener As clsCmbListener


Private Sub Form_Load()
    
    'Const
    Const FUNC_NAME As String = "Form_Load"
    
    'Variable
    Dim dicInfo As Object
    
    On Error GoTo ErrorHandler

    'set Event Class
    Set objCmbListener = New clsCmbListener: Set objCmbListener.ComboBox = Me.cmb_withEventsTest
    
    'set Event Log
    Set M_EventLog.targetTxtBox = Me.txt_EventLog
    
ExitHandler:

    Exit Sub
    
ErrorHandler:

    MsgBox "An error has occurred and the macro will be terminated." & _
           vbLf & _
           "Func Name:" & FUNC_NAME & _
           vbLf & _
           "Error No." & Err.Number & vbNewLine & _
           Err.Description, vbCritical, "Access-Control-WithEvents"
        
    GoTo ExitHandler
        
End Sub


Private Sub cmb_withEventsTest_BeforeUpdate(Cancel As Integer)
    
    'Const
    Const FUNC_NAME As String = "cmb_withEventsTest_BeforeUpdate"
    
    'Variable
    
    On Error GoTo ErrorHandler

    'do logging
    If Not M_EventLog.writeEventLogs(FUNC_NAME) Then GoTo ExitHandler

ExitHandler:

    Exit Sub
    
ErrorHandler:

    MsgBox "An error has occurred and the macro will be terminated." & _
           vbLf & _
           "Func Name:" & FUNC_NAME & _
           vbLf & _
           "Error No." & Err.Number & vbNewLine & _
           Err.Description, vbCritical, "Access-Control-WithEvents"
        
    GoTo ExitHandler
        
End Sub

Private Sub cmb_withEventsTest_AfterUpdate()
    
    'Const
    Const FUNC_NAME As String = "cmb_selectedRcd_AfterUpdate"
    
    'Variable
    
    On Error GoTo ErrorHandler
    
    'do logging
    If Not M_EventLog.writeEventLogs(FUNC_NAME) Then GoTo ExitHandler
    If Not M_EventLog.writeEventLogs("""" & Me.cmb_withEventsTest.Value & """" & "Selected") Then GoTo ExitHandler

ExitHandler:

    Exit Sub
    
ErrorHandler:

    MsgBox "An error has occurred and the macro will be terminated." & _
           vbLf & _
           "Func Name:" & FUNC_NAME & _
           vbLf & _
           "Error No." & Err.Number & vbNewLine & _
           Err.Description, vbCritical, "Access-Control-WithEvents"
        
    GoTo ExitHandler
        
End Sub
```



<br><br>

#### COMBOBOX EVENT LISTENER CLASS

```vb
Option Compare Database
Option Explicit

'**************************
'*Combobox Event Listener
'**************************

'Const

'Variable
Private WithEvents myComboBox As Access.ComboBox

'******************************************************************************************
'*getter/setter
'******************************************************************************************
Public Property Set ComboBox(ByRef cmb As Access.ComboBox)
    Set myComboBox = cmb
    myComboBox.OnKeyDown = "[Event Procedure]"
End Property





'******************************************************************************************
'*Function ：disable keyboard input
'*Arg(1)   ：key code
'*Arg(2)   ：shft key pressed or not
'******************************************************************************************
Private Sub myComboBox_KeyDown(KeyCode As Integer, Shift As Integer)
    
    'Const
    Const FUNC_NAME As String = "myComboBox_KeyDown"
    
    'Variable
    
    On Error GoTo ErrorHandler
    
    'prohibit entering (except fror Enter/Tab/Esc)
    If KeyCode = vbKeyReturn Then GoTo ExitHandler
    If KeyCode = vbKeyTab Then GoTo ExitHandler
    If KeyCode = vbKeyEscape Then GoTo ExitHandler
    
    KeyCode = 0
    
    If Not M_EventLog.writeEventLogs(FUNC_NAME) Then GoTo ExitHandler
    
ExitHandler:

    Exit Sub
    
ErrorHandler:

    MsgBox "An error has occurred and the macro will be terminated." & _
           vbLf & _
           "Func Name:" & FUNC_NAME & _
           vbLf & _
           "Error No." & Err.Number & vbNewLine & _
           Err.Description, vbCritical, "Access-Control-WithEvents"
        
    GoTo ExitHandler
        
End Sub

```

<br><br>


#### EVENT LOG MODULE

```vb
Option Compare Database
Option Explicit


'**************************
'*Event Log Module
'**************************

'Const


'Variable
Public targetTxtBox As Access.TextBox


'******************************************************************************************
'*Function ：write the event log into the textbox specified in a module variable
'*Arg(1)   ：the written string
'*Return   ：True > normal termination; False > abnormal termination

'******************************************************************************************
Public Function writeEventLogs(ByVal logTxt As String) As Boolean
    
    'Const
    Const FUNC_NAME As String = "writeEventLogs"
    
    'Variable
    
    On Error GoTo ErrorHandler

    writeEventLogs = False
    
    If Nz(targetTxtBox.Value, "") <> "" Then targetTxtBox.Value = targetTxtBox.Value & vbNewLine
    targetTxtBox.Value = targetTxtBox.Value & _
                         Now & _
                         " : " & _
                         logTxt
    
    writeEventLogs = True
    
ExitHandler:

    Exit Function
    
ErrorHandler:

    MsgBox "An error has occurred and the macro will be terminated." & _
           vbLf & _
           "Func Name:" & FUNC_NAME & _
           vbLf & _
           "Error No." & Err.Number & vbNewLine & _
           Err.Description, vbCritical, "Access-Control-WithEvents"
        
    GoTo ExitHandler
        
End Function

```

## TROUBLE

The combobox KeyDown event is Supposed to be captured by `myComboBox_KeyDown` function in `clsCmbListener` class through the mechanism of `WithEvents`,  
and Keystrokes should be prohibited except for some keys such as Enter and Tab.

However, I saw the combobox allowing manual input.  
{{< video src="https://res.cloudinary.com/ddxhi1rnh/video/upload/v1644740250/learnerBlog/Access-Control-WithEvents/en/Media1_pkuc04.webm" max_width=600px is_bundle=false >}}

Furthermore, since the log which indicates `myComboBox_KeyDown` has called is not displayed in the textbox, `WithEvents` is not working as expected in the first place.

In the case of MS Excel, above went well. So I pondered what to do for a while.



## THE SOLUTION

### i. ADD KEYDOWN EVENT FUNCTION INTO THE FORM MODULE AS WELL

#### EXPLANATION

I added a function whose processing is empty.


```vb
Option Compare Database
Option Explicit

'**************************
'*MainForm
'**************************

'Const


'Variable
Private objCmbListener As clsCmbListener


Private Sub cmb_withEventsTest_KeyDown(KeyCode As Integer, Shift As Integer)
'empty
End Sub
```

By doing so, the KeyDown event turned to be called.  
{{< video src="https://res.cloudinary.com/ddxhi1rnh/video/upload/v1644740250/learnerBlog/Access-Control-WithEvents/en/Media2_hqxppr.webm" max_width=600px is_bundle=false >}}




#### ONE FURTHER PROBLEM

However, this way contains one further problem.

When above `cmb_withEventsTest_KeyDown` function is truly empty, `VBE` mechanism automatically delete the function during compiling phase because it's not necessary.

So, it has to have one comment row and escape the deletion, that makes the tool less maintainable.  
Moreover, when others see this code, they might consider it unnecessary and delete it.



### ii. SET [EVENT PROCEDURE] TO THE ONKEYDOWN PROPERTY OF THE COMBOBOX INSTANCE

The solution is taken from [this stackoverflow](https://stackoverflow.com/questions/23522230/creating-a-class-to-handle-access-form-control-events)

As above articel says, `[Event Procedure]` is the key of the solving the problem.

```vb
listener.ct.OnClick = "[Event Procedure]"  '<------- Assigned the event handler
```

I applied this logic to my own code.

```vb
'**************************
'*Combobox Event Listener
'**************************

'******************************************************************************************
'*getter/setter
'******************************************************************************************
Public Property Set ComboBox(ByRef cmb As Access.ComboBox)
    Set myComboBox = cmb
    myComboBox.OnKeyDown = "[Event Procedure]"
End Property


```

After that, my code started to work fine!



