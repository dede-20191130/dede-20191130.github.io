---
title: 'TOOL DEVELOPMENT ON EXCEL: COMPARING BETWEEN TOOLS USING FUNCTION (NON-MACRO) AND TOOLS USING VBA MACRO'
author: dede-20191130
date: 2020-07-16T06:14:50+00:00
slug: cmpr-tools
draft: false
tags: ['Excel', 'VBA',  'Excel Function']
categories: ['Application', 'programming']
vba_taxo: general
toc: true
archives:
    - 2020
    - 2020-07
---

{{< help-me-improve-lang  >}}

## ABOUT THIS ARTICLE

Let's assume you want to create tool by Excel that enables us to manage data or output some documents.  
Broadly speaking, the following options may be considered.  

* tool using Excel build-in function, not using VBA macro.
* tool using VBA macro.

Each option has its advantages and disadvantages.  
I'd like to summarize them from my personal point of view.



## SUMMARY TABLE 

First of all, I wrote summary table.

<div class="">
    <div class="">
        <table class="">
            <tr>
                <td>
                </td>
                <td>
                    Tool using Excel func
                </td>
                <td>
                    Tool using VBA macro
                </td>
            </tr>
            <tr>
                <td>
                    Difficulty of development
                </td>
                <td>
                    Low
                </td>
                <td>
                    High
                </td>
            </tr>
            <tr>
                <td>
                    Degree of Freedom
                </td>
                <td>
                    Low
                </td>
                <td>
                    High
                </td>
            </tr>
            <tr style="vertical-align: top;">
                <td>
                    Things we can do
                </td>
                <td>
                    <ul>
                        <li>Data management<br/> (manually input, working with databases)</li>
                        <li>Output products, for example ledger sheet
                            <ul>
                                <li>Print by printer</li>
                                <li>Output in PDF format</li>
                            </ul>
                        </li>
                        <li>Mail Creation</li>
                        <li>File-path management</li>
                        <li>Web-sites-urls management</li>
                    </ul>
                </td>
                <td>
                    <ul>
                        <li>Data management
                            <ul>
                                <li>manually input</li>
                                <li>input on original input form</li>
                                <li>file reading, writing</li>
                                <li>working with databases</li>
                            </ul>
                        </li>
                        <li>Output products, for example ledger sheet
                            <ul>
                                <li>Print by printer</li>
                                <li>Output in PDF format</li>
                                <li>CSV File</li>
                                <li>Any other format can be output</li>
                            </ul>
                        </li>
                        <li>
                            The following items at a more advanced level
                            <ul>
                                <li>Mail Creation</li>
                                <li>File-path management</li>
                                <li>Web-sites-urls management</li>
                            </ul>
                        </li>
                        <li>Scraping web-sites</li>
                        <li>Request and Response for web-sites</li>
                    </ul>
                </td>
            </tr>
            <tr>
                <td>
                    Degree of improvement in the accuracy of the work
                </td>
                <td>
                    Low
                </td>
                <td>
                    High
                </td>
            </tr>
            <tr>
                <td>
                    Degree of improvement in work efficiency
                </td>
                <td>
                    Low
                </td>
                <td>
                    High
                </td>
            </tr>
            <tr>
                <td>
                    Stability of operation
                </td>
                <td>
                    High
                </td>
                <td>
                    Low
                </td>
            </tr>
            <tr>
                <td>
                    Maintainability
                </td>
                <td>
                    High
                </td>
                <td>
                    Low
                </td>
            </tr>
        </table>
    </div>
</div>



## COMPARISON

### EASE OF CREATING

A tool using only Excel build-in function is easy of creating with basic knowledge of mathematical function.  
When you are using IF function, VLOOKUP function, or OFFSET function, you need idea and knowledge of programming.

On the other hand, for creating a tool using VBA macro, programming work using Visual Basic is required.  
So it's hard for inexperienced person of programming.


### DEGREE OF FREEDOM AND THINGS WE CAN DO

What Excel functions can do is limited to data management, formatting, input/output, document creation, output, and emailing etc.  
It can't manipulate outer application.

For VBA macro, we can manipulate text file, binary file, of course MS Office Application.  
You can even emulate keystrokes with the Sendkey function to control other applications.  
When manipulating browser, you use Selenium.

There is so much you can do with VBA macro, but the more complex you try to make it, the more bugs and instability you will encounter.

### BENEFITS OF AUTOMATION

VBA macro tool, having high degree of freedom and ability of performing a series of tasks quickly and with precision, is maximize benefits of automation.

### STABILITY

The behavior of Excel functions is predetermined, so there is little lisk of bug in a processing assembled by Excel functions.  
In addition to this, when using Excel function and error happening, the error will be displayed as "#VALUE!" in the cell where the function is entered.  
So you can easily detect the error.

For VBA macro tool, depending on how to write the program, it may force Excel to close itself or contaminate the data in the text file on being edited.

For example, during editing text file, if Excel crashes with an error, What you don't expect may be written in text file.

### MAINTAINANCE

For maintaining vba macro, a person who can read vb code is needed.  
So, in case that that person retires or moves, Nobody remaining may not be able to maintain the tool.

Therefore, it's important to create even simple documentation for its internal Specifications.


## SUMMARY

Non-macro tool has high maintainability and stability, but less benefit of automation.  
When you created a tool using VBA macro, you'd better keep the documentation of the specifications.
