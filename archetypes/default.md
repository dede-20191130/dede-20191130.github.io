---
title: "{{ replace .Name "-" " " | title }}"
author: dede-20191130
date: {{ .Date }}
slug: {{ .Name }}
draft: true
toc: true
featured: false
tags: []
categories: []
archives:
    - {{ now.Format "2006" }}
    - {{ now.Format "2006-01" }}
shareImage: "images/thumbnail.png" 
---

