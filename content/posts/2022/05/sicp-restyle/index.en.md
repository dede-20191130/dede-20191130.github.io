---
title: "SCRIPT FOR SCRIPT_AUTO_RUNNER TO MAKE SICP TEXTBOOK MORE READABLE"
author: dede-20191130
date: 2022-05-22T20:03:42+09:00
slug: sicp-restyle
draft: false
toc: true
featured: false
tags: ["javascript","Chrome"]
categories: ["programming"]
archives:
    - 2022
    - 2022-05
---


{{< help-me-improve-lang  >}}

## ABOUT THIS ARTICLE

{{< box-with-title title="Brief Summary" >}} 
    We can do syntax-highlighting to unhighlighted code-blocks using a library from CDN.<br/>
    We use ScriptAutoRunner for execution environment.

{{< /box-with-title >}}

Hello, I'm dede.

In this article, I'll introduce a way to make SICP textbook, which is available to the public on the web, more readable  
by using a short script. 


## WHAT'S SICP
### ABOUT

SICP textbook is one of the most famous and effective CS books.  
It describes frequent patterns in computer programming, such as recursive procedures and the benefits of modularity. 


### PUBLISHED WEB SITE

[Here](https://mitpress.mit.edu/sites/default/files/sicp/full-text/book/book.html) is a published content in web site equivalent to its book version.


### SOME PROBREMS

About above web site I found myself a bit difficult to read in several respects.

- The book uses _Scheme_, one of the variations of lisp language, but code blocks are not syntax-highlighted so are not very readable.
- The texts in footnote are a bit small.

I wrote a JS script to modify their design which runs on browser.

Normally in order to run self-made Javascript code on browser, we take a way either to run it interactively in developer-console or to register the code as a Bookmarklet and run it.   
But both of them have a disadvantage that we must run it whenever we reload the page.

However there is a useful extension with respect to Chrome.




## WHAT'S SCRIPT_AUTO_RUNNER

[[ScriptAutoRunner](https://chrome.google.com/webstore/detail/scriptautorunner/gpgjofmpmjjopcogjgdldidobhmjmdbm?hl=ja)

This is a useful Chrome extension in which we register a pair of specific domain string and script code we want to run, and whenever we open/reload a page of target domain the code runs automatically.

![ScriptAutoRunner from official](https://lh3.googleusercontent.com/LUHrciH1gr-dNe_0yrVuje-TYIb66LIJePum2HDipQ8HFPB_kjpvQqLnYxbw7Wn_drDTLf7l604zciVYugAUvg6ic00=w640-h400-e365-rj-sc0x00ffffff)

## CREATION ENVIRONMENT AND USED TOOL

- Google Chrome version: 101.0.4951.67
- highlight.js v11.5.1

## SCRIPT CODE
### WHOLE CODE

Here is a whole code.  
Following sections describe meaning of each and its behavior.


```js
// if document is not fully loaded, load event take the place
if (document.readyState === "complete") {
    restyleSICP();
} else {
    window.addEventListener("load", restyleSICP);
}

function restyleSICP() {
    // only in specific paths this script runs
    if (!window.location.pathname.startsWith("/sites/default/files/sicp/full-text/book")) return;
    highlightSchemeCode();
    expandFootnote();
}

function highlightSchemeCode() {
    // exclude elements which contain some img elements
    const targetTts = Array.from(document.querySelectorAll("p > tt:first-child:last-child")).filter(e => !e.querySelector("img"));

    for (const tt of targetTts) {
        const pre = document.createElement("pre");
        const code = document.createElement("code");

        code.classList.add("language-scheme");
        code.innerHTML = tt.innerHTML;

        // replace tt element with pre + code elements
        pre.append(code);
        tt.before(pre);
        tt.remove();
    }

    const link = document.createElement("link");
    link.rel = "stylesheet";
    link.href = "https://cdnjs.cloudflare.com/ajax/libs/highlight.js/11.5.1/styles/a11y-dark.min.css";

    const hlScr = document.createElement("script");
    hlScr.src = "//cdnjs.cloudflare.com/ajax/libs/highlight.js/11.5.1/highlight.min.js";
    hlScr.onload = () => {
        hljs.configure({
            ignoreUnescapedHTML: true
        });
        const schemeScr = document.createElement("script");
        schemeScr.src = "https://cdnjs.cloudflare.com/ajax/libs/highlight.js/11.5.1/languages/scheme.min.js";
        schemeScr.onload = () => hljs.highlightAll();
        document.head.append(schemeScr);

    }
    document.head.append(link, hlScr);


}

function expandFootnote() {
    document.querySelector(".footnote").style.fontSize = "1rem";
}
```

### LAUNCHING

```js
// if document is not fully loaded, load event take the place
if (document.readyState === "complete") {
    restyleSICP();
} else {
    window.addEventListener("load", restyleSICP);
}
```

The running timing of main processing is diferrent between the case in which DOM reading of the page is not yet complete and the case in which it has already been complete.



### CHECK IF RUN OR NOT

```js
function restyleSICP() {
    // only in specific paths this script runs
    if (!window.location.pathname.startsWith("/sites/default/files/sicp/full-text/book")) return;
    // ...
}
```

Unfortunately ScriptAutoRunner allows us to specify a domain but not a path, so we check if it can run Subsequent processing or not by fetching the page's url.



### HIGHLIGHTING

```js
function highlightSchemeCode() {
    // exclude elements which contain some img elements
    const targetTts = Array.from(document.querySelectorAll("p > tt:first-child:last-child")).filter(e => !e.querySelector("img"));

    for (const tt of targetTts) {
        const pre = document.createElement("pre");
        const code = document.createElement("code");

        code.classList.add("language-scheme");
        code.innerHTML = tt.innerHTML;

        // replace tt element with pre + code elements
        pre.append(code);
        tt.before(pre);
        tt.remove();
    }

    
```

All Page's Elements corresponding to _Scheme_ code block are _tt_ elements.  
We search them and replace _pre_ element + _code_ element so that _highlight.js_ finds them as a highlighting target.

Some code blocks contains _img_ elementsm and We exclude them because highlighting is not effective for them.



```js
const link = document.createElement("link");
    link.rel = "stylesheet";
    link.href = "https://cdnjs.cloudflare.com/ajax/libs/highlight.js/11.5.1/styles/a11y-dark.min.css";

    const hlScr = document.createElement("script");
    hlScr.src = "//cdnjs.cloudflare.com/ajax/libs/highlight.js/11.5.1/highlight.min.js";
    hlScr.onload = () => {
        hljs.configure({
            ignoreUnescapedHTML: true
        });
        const schemeScr = document.createElement("script");
        schemeScr.src = "https://cdnjs.cloudflare.com/ajax/libs/highlight.js/11.5.1/languages/scheme.min.js";
        schemeScr.onload = () => hljs.highlightAll();
        document.head.append(schemeScr);

    }
    document.head.append(link, hlScr);


}
```

Via CDN we read main script of hl and optional script for _Scheme_ successively and when all completed `highlightAll` runs.



### RESTYLE A FOOTNOTE

```js
function expandFootnote() {
    document.querySelector(".footnote").style.fontSize = "1rem";
}
```

## DEMO

![before](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1653230225/learnerBlog/sicp-restyle/mitpress.mit.edu_sites_default_files_sicp_full-text_book_book-Z-H-12.html_aiwgpd.png)

![after](https://res.cloudinary.com/ddxhi1rnh/image/upload/v1653230225/learnerBlog/sicp-restyle/mitpress.mit.edu_sites_default_files_sicp_full-text_book_book-Z-H-12.html_1_demakx.png)

