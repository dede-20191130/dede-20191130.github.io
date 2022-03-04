---
title: "ACCESS VBA: I CREATED A FUNCTION EXPORTING TABLE DEFINITIONS DISPLAYED AT DESIGN VIEW IN A TABULAR FORMAT."
author: dede-20191130
date: 2020-10-25T22:01:43+09:00
slug: Access-tableDef
draft: False
toc: true
tags: ['Access', 'VBA','HOMEMADE']
categories: ['problem-solving', 'programming']
vba_taxo: help_develop
archives:
    - 2020
    - 2020-10
---

{{< help-me-improve-lang  >}}

## ABOUT THIS ARTICLE

The options for exporting each table in MSAccess includes:  
- Push export from the right-click menu of table item in Navigation Bar -> Select the format such as Excel or text and export it.
- Push Database Documenter in Database Tools tab -> Specify target table -> Select the format such as Print, Excel or PDF and export it.

However, the former can't export with the detailed table information such as field types and the existence of primary keys,  
and the latter can export with the detailed information, but the data format is Single Form for each field,   
so it's difficult to grasp the table setting like in the design view.

![Table Information As a List In Design View](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1644651644/learnerBlog/Access-tableDef/en/%E3%82%B9%E3%82%AF%E3%83%AA%E3%83%BC%E3%83%B3%E3%82%B7%E3%83%A7%E3%83%83%E3%83%88_2022-02-12_164024_w8ct3k.png) 

So, existing features cannot meet the demand, I created a function exporting table definitions displayed at design view in a tabular format.


{{< colored-span color="#fb9700" >}}I created a tool version of this functionality.🔽🔽 {{< /colored-span >}}  
{{< page-titled-link page="access-tabledef-tool" >}}




## CREATION ENVIRONMENT
Microsoft Office 2019

## FUNCTION OVERVIEW

Export all table definitions as a tabular format in the current project.  
The export destination is new Excel Book. All table meta data is stored in separate sheets.



## FUNCTIONS

|NAME|KIND|FUNCTIONALITY|
| ---- | ---- | ---- |
|exportTableDefTablesMain|Sub Procedure|main function for export|
|getTableDefArray|Function Procedure|get definition information of the table |
|getFieldTypeString|Function Procedure|get field type string of argument field|
|getPKs|Function Procedure|get field strings of primary keys|
|getFKs|Function Procedure|get field strings of foreign keys|
|setWSName|Function Procedure|set worksheet Name to argument sheet|

  

## CALLER-CALLEE RELATION
  
- exportTableDefTablesMain
    - calling -> getTableDefArray
        - calling -> getFieldTypeString
        - calling -> getPKs
        - calling -> getFKs
    - calling -> setWSName




  
## CODES

### [exportTableDefTablesMain]
```vb
'******************************************************************************************
'*Function      :export data and create excel book
'******************************************************************************************
Public Sub exportTableDefTablesMain()
    
    'Const
    Const FUNC_NAME As String = "exportTableDefTablesMain"
    
    'Variable
    Dim xlApp As Object
    Dim wb As Object
    Dim tdf As DAO.TableDef
    Dim defArr As Variant
    Dim fstWs As Object
    Dim ws As Object
    
    On Error GoTo ErrorHandler
    
    'create new excel app instance and excel-book instance
    Set xlApp = CreateObject("Excel.Application")
    With xlApp
        .Visible = False
        .DisplayAlerts = False
        .ScreenUpdating = False
    End With
    Set wb = xlApp.Workbooks.Add
    
    Set fstWs = wb.Worksheets(1)
    
    'create a access table definition information table in separate sheets
    For Each tdf In CurrentDb.TableDefs
        Do
            'do continue if tdf is one of the unnecessary tables such as system table.
            If Left(tdf.Name, 4) = "Msys" Or Left(tdf.Name, 4) = "Usys" Or Left(tdf.Name, 1) = "~" Then Exit Do
            
            'get definition information of the table 
            defArr = getTableDefArray(tdf)
            If IsNull(defArr) Then GoTo ExitHandler
            
            'create a new sheet
            Set ws = wb.Worksheets.Add
            If Not setWSName(ws, tdf.Name) Then Call Err.Raise(1000, "Sheet Name Specification Error", "An error has occurred on sheet name specification.")
            
            'write a definition information to Range and auto-adjust the sheet column widths.
            With ws.Range(ws.cells(1, 1), ws.cells(UBound(defArr) - LBound(defArr) + 1, UBound(defArr, 2) - LBound(defArr, 2) + 1))
                .Value = defArr
                .EntireColumn.AutoFit
            End With
            
        Loop While False
    Next tdf
    
    'remove default sheet
    If wb.Worksheets.Count > 1 Then fstWs.Delete
    
    'save
    wb.saveas Application.CurrentProject.Path & _
              "\" & _
              Left( _
              CurrentProject.Name, _
              InStrRev(CurrentProject.Name, ".") - 1 _
              ) & _
                "_Table_Info_List.xlsx"
    
    'Complete
    MsgBox "Completed", , "INFO"
    
    
ExitHandler:
    
    'close connection and instances
    If Not wb Is Nothing Then wb.Close: Set wb = Nothing
    If Not xlApp Is Nothing Then xlApp.Quit: Set xlApp = Nothing
    
    Set tdf = Nothing
    Set ws = Nothing
    Set fstWs = Nothing
    
    Exit Sub
    
ErrorHandler:

    MsgBox "An error has occurred and the macro will be terminated." & _
           vbLf & _
           "Func Name:" & FUNC_NAME & _
           vbLf & _
           "Error No." & Err.Number & vbNewLine & _
           Err.Description, vbCritical, "Error"

        
    GoTo ExitHandler
        
End Sub
```
Launch the Excel application as a new instance.  
A default sheet will be removed after storing data.



  
    
  

### [getTableDefArray]
```vb
'******************************************************************************************
'*Function      : get definition information of the table 
'*                Items: 
'*                      Field Name
'*                      Data Type
'*                      Size
'*                      Required or not
'*                      Primary key or not
'*                      Foreign key or not
'*                      Description
'*arg(1)        : TableDef instance
'*return        : definition information array
'******************************************************************************************

Public Function getTableDefArray( _
       ByVal pTdf As DAO.TableDef _
       ) As Variant
    
    'Const
    Const FUNC_NAME As String = "getTableDefArray"
    
    'Variable
    Dim defArr() As Variant
    Dim fld As DAO.Field
    Dim i As Long
    Dim dicPKs As Object
    Dim dicFKs As Object
    Dim description As String
    
    On Error GoTo ErrorHandler

    getTableDefArray = Null
    
    'do redimension the array to be (Field Count + 1) rows and 7 columns
    ReDim defArr(0 To pTdf.Fields.Count, 0 To 6)
    
    'setting the header part
    defArr(0, 0) = "Field Name"
    defArr(0, 1) = "Data Type"
    defArr(0, 2) = "Size"
    defArr(0, 3) = "Required or not"
    defArr(0, 4) = "Primary key or not"
    defArr(0, 5) = "Foreign key or not"
    defArr(0, 6) = "Description"
    
    'get a dictionary containing all primary key field names
    Set dicPKs = getPKs(pTdf)
    If dicPKs Is Nothing Then GoTo ExitHandler
    
    'get a dictionary containing all foreign key field names
    Set dicFKs = getFKs(pTdf)
    If dicFKs Is Nothing Then GoTo ExitHandler
    
    For i = 1 To pTdf.Fields.Count
        Set fld = pTdf.Fields(i - 1)
        'Field Name
        defArr(i, 0) = fld.Name
        'Data Type
        defArr(i, 1) = getFieldTypeString(fld.Type)
        'Size
        If fld.Type = dbText Then
            defArr(i, 2) = fld.Size
        Else
            defArr(i, 2) = "-"
        End If
        'Required or not
        If fld.Required Then defArr(i, 3) = ChrW("&H" & 2714)
        'Primary key or not ◆note1
        If dicPKs.Exists(fld.Name) Then defArr(i, 4) = ChrW("&H" & 2714)
        'Foreign key or not ◆note1
        If dicFKs.Exists(fld.Name) Then defArr(i, 5) = ChrW("&H" & 2714)
        'Description
        On Error Resume Next
        description = fld.Properties("Description")
        On Error GoTo ErrorHandler
        defArr(i, 6) = description
    Next i


    getTableDefArray = defArr
    
ExitHandler:
    
    Set fld = Nothing
    Set dicFKs = Nothing
    Set dicPKs = Nothing
    
    Exit Function
    
ErrorHandler:

    MsgBox "An error has occurred and the macro will be terminated." & _
           vbLf & _
           "Func Name:" & FUNC_NAME & _
           vbLf & _
           "Error No." & Err.Number & vbNewLine & _
           Err.Description, vbCritical, "Error"

    GoTo ExitHandler
        
End Function
```

◆note1: With dicPKs and dicFKs, check if current field in loop is included in the dics.



### [getFieldTypeString]
```vb
'******************************************************************************************
'*Function      : get field type string of argument field
'*arg(1)        : field type number
'*return        : field data string
'******************************************************************************************
Public Function getFieldTypeString(ByVal pFldTyepNum As Long) As String
    
    'Const
    Const FUNC_NAME As String = "getFieldTypeString"
    
    'Variable
    Dim strType As String
    
    On Error GoTo ErrorHandler

    strType = ""
    

    Select Case pFldTyepNum
    Case dbBoolean
        strType = "Bool"
    Case dbByte
        strType = "Byte"
    Case dbInteger
        strType = "Integer"
    Case dbLong
        strType = "Long Integer"
    Case dbSingle
        strType = "Single Number"
    Case dbDouble
        strType = "Double Number"
    Case dbCurrency
        strType = "Currency"
    Case dbDate
        strType = "Date"
    Case dbText
        strType = "short Text"
    Case dbLongBinary
        strType = "OLE Object Type"
    Case dbMemo
        strType = "long text"
    End Select

    getFieldTypeString = strType
    
ExitHandler:

    Exit Function
    
ErrorHandler:

    MsgBox "An error has occurred and the macro will be terminated." & _
           vbLf & _
           "Func Name:" & FUNC_NAME & _
           vbLf & _
           "Error No." & Err.Number & vbNewLine & _
           Err.Description, vbCritical, "Error"

        
    GoTo ExitHandler
        
End Function
```

Each data type of DAO.Field Object Type property is number, So the function converts it to string.



### [getPKs]
```vb
'******************************************************************************************
'*Function      : get field strings of primary keys
'*arg(1)        : TableDef instance
'*return        : Dictionary containing PK info
'******************************************************************************************

Public Function getPKs(ByVal pTdf As DAO.TableDef) As Object
    
    'Const
    Const FUNC_NAME As String = "getPKs"
    
    'Variable
    Dim idx As DAO.Index
    Dim fld As DAO.Field
    Dim dic As Object
    
    On Error GoTo ErrorHandler

    Set getPKs = Nothing
    Set dic = CreateObject("Scripting.Dictionary")
    
    
    'check if Primary property is true
    For Each idx In pTdf.Indexes
        If idx.Primary = True Then
            For Each fld In idx.Fields
                dic.Add fld.Name, True
            Next
        End If
    Next

    'Return
    Set getPKs = dic
    
ExitHandler:

    Set dic = Nothing

    Exit Function
    
ErrorHandler:

    MsgBox "An error has occurred and the macro will be terminated." & _
           vbLf & _
           "Func Name:" & FUNC_NAME & _
           vbLf & _
           "Error No." & Err.Number & vbNewLine & _
           Err.Description, vbCritical, "Error"

    GoTo ExitHandler
        
End Function
```

### [getFKs]
```vb
'******************************************************************************************
'*Function      : get field strings of foreign keys
'*arg(1)        : TableDef instance
'*return        : Dictionary containing PK info
'******************************************************************************************
Public Function getFKs(ByVal pTdf As DAO.TableDef) As Object
    
    'Const
    Const FUNC_NAME As String = "getFKs"
    
    'Variable
    Dim rsRelation As DAO.Recordset
    Dim dic As Object
    
    On Error GoTo ErrorHandler

    Set getFKs = Nothing
    Set dic = CreateObject("Scripting.Dictionary")
    
    'access MSysRelationships system table
    Set rsRelation = CurrentDb.OpenRecordset( _
                     "SELECT szColumn FROM MSysRelationships WHERE szObject =" & _
                     " " & _
                     "'" & _
                     pTdf.Name & _
                     "'" & _
                     ";" _
                     )
    
    With rsRelation
        If .EOF Then Set getFKs = dic: GoTo ExitHandler
        .MoveFirst
        Do Until .EOF
            dic.Add .Fields("szColumn").Value, True
            .MoveNext
        Loop
    End With
    
    'Return
    Set getFKs = dic
    
ExitHandler:
    
    If Not rsRelation Is Nothing Then rsRelation.Close: Set rsRelation = Nothing
        
    Exit Function
    
ErrorHandler:

    MsgBox "An error has occurred and the macro will be terminated." & _
           vbLf & _
           "Func Name:" & FUNC_NAME & _
           vbLf & _
           "Error No." & Err.Number & vbNewLine & _
           Err.Description, vbCritical, "Error"
        
    GoTo ExitHandler
        
End Function
```

Foreign keys information is stored in a system table, so the function access it.





### [setWSName]
```vb
'******************************************************************************************
'*Function      : set worksheet Name to argument sheet
'*arg(1)        : excel worksheet instance
'*arg(2)        : the name set
'*return        : True > normal termination; False > abnormal termination
'******************************************************************************************
Public Function setWSName( _
       ByVal ws As Object, _
       ByVal newName As String _
       ) As Boolean
    
    'Const
    Const FUNC_NAME As String = "setWSName"
    
    'Variable
    
    On Error GoTo ErrorHandler

    setWSName = False
    
    ws.Name = newName

    setWSName = True
    
ExitHandler:

    Exit Function
    
ErrorHandler:

    'escaping route: if the name includes some charactors not allowed to use to sheet name
    ws.Name = "Table_" & ws.Parent.Worksheets.Count & "_" & Format(Now, "yyyymmddhhnnss")

    setWSName = True
    GoTo ExitHandler
        
End Function
```






## HOW TO USE

Write above functions somewhere in the module of the Access file you want to extract table info,   
and run `exportTableDefTablesMain`.



## REFERENCE IMAGES

An Excel file will be output as shown below.  
You can overview the definition information in a style that resembles the Design View display.





![Export Sample No1](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1644651504/learnerBlog/Access-tableDef/en/Access-tableDef_q1qoik.png)  
![Export Sample No2](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1644651504/learnerBlog/Access-tableDef/en/Access-tableDef2_rhocp4.png)