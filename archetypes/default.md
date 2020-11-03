---
title: "{{ replace .Name "-" " " | title }}"
author: dede-20191130
date: {{ .Date }}
slug: foobar
draft: true
toc: true
tags: []
categories: []
archives:
    - {{ now.Format "2006" }}
    - {{ now.Format "2006-01" }}
---

