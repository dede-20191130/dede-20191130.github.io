---
title: "ACCESS VBA: I CREATED A TOOL EXPORTING TABLE DEFINITIONS DISPLAYED AT DESIGN VIEW IN A TABULAR FORMAT."
author: dede-20191130
date: 2020-10-27T20:29:45+09:00
slug: Access-TableDef-Tool
draft: false
toc: true
tags: ['Access', 'VBA','Homebrew','tool']
categories: ['application', 'programming']
vba_taxo: help_develop
archives:
    - 2020
    - 2020-10
---

{{< help-me-improve-lang  >}}

## ABOUT THIS ARTICLE

This is a tool version of the export function described in the [previous article]({{< relref "access-tabledef.md" >}}).  

By toolizing, it's no longer needed to open VBE and paste the functions.  
And it enabled us to export definition information of other access database file.

You can download the tool and view its source code from [here](https://github.com/dede-20191130/My_VBA_Tools/tree/master/T0002_Access%E3%83%86%E3%83%BC%E3%83%96%E3%83%AB%E5%AE%9A%E7%BE%A9%E3%82%A8%E3%82%AF%E3%82%B9%E3%83%9D%E3%83%BC%E3%83%88%E3%83%84%E3%83%BC%E3%83%AB/en)!

  
## CREATION ENVIRONMENT
Microsoft Office 2019



## FUNCTION OVERVIEW

Ability to export table definitions in tabular format





## ADVANTAGE OF THE TOOL

- It makes it easy to list table definitions and visually assess the information.
- It gives you the approach to the information more efficiently than creating an equivalent table manually.







## HOW TO USE

### 1. SELECT ACCESS FILE PATH TO BE EXPORTED

![Export Screen](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1644677176/learnerBlog/Access-TableDef-Tool/en/%E3%82%B9%E3%82%AF%E3%83%AA%E3%83%BC%E3%83%B3%E3%82%B7%E3%83%A7%E3%83%83%E3%83%88_2022-02-12_234426_ul5wq0.png)

Write a file path by manually input or via dialog box displayed after pushing Browse button.



### 2. RUN EXPORTING

Push the Export button.

An Excel file is output to a folder on the same level as the target Access file as shown below.  
With the file, you can overview the table definition information in a style that resembles the Design View display. 

![Export Sample No1](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1644677437/learnerBlog/Access-TableDef-Tool/en/Access-tableDef-tool01_qxvwb5.png)
![Export Sample No2](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1644677437/learnerBlog/Access-TableDef-Tool/en/Access-tableDef-tool02_d1ubn8.png)



