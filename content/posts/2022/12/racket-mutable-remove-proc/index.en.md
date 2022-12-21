---
title: "RACKET REMOVE PROCEDURE FOR MUTABLE LISTS"
author: dede-20191130
date: 2022-12-21T15:10:56+09:00
slug: racket-mutable-remove-proc
draft: false
toc: true
featured: false
tags: ["racket"]
categories: ["programming"]
archives:
    - 2022
    - 2022-12
---

## ABOUT THIS ARTICLE

Racket's buildin library has remove procedure for immutable lists [[link]](https://docs.racket-lang.org/reference/pairs.html#%28def._%28%28lib._racket%2Fprivate%2Flist..rkt%29._remove%29%29).  
But the one for mutable lists is missing, so I created and test with some codes.

## ENVIRONMENT

Operation has been tested in Racket v8.5.

## SAMPLES

<script src="https://gist.github.com/dede-20191130/7166dd592d02bb3e2966add1574e0072.js"></script>