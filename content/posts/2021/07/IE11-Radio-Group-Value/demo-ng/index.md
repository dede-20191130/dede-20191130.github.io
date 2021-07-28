---
title: "[Javascript] IE11でRadioボタングループの値を取得できない事象＆対策 NGデモ"
author: dede-20191130
date: 2021-07-28T19:14:46+09:00
slug: IE11-Radio-Group-Value-demo-ng
draft: false
toc: true
featured: false
tags: []
categories: []
archives:
    - 2021
    - 2021-07
type: _default
layout: plain
---

<main>
    <div id="container">
        <form id="my-form" name="my-form" onsubmit="return false;">
            <h3>県庁所在地の表示</h3>
            <div>
                <input id="radio01-toyama" type="radio" name="radio01" value="富山市">
                <label for="radio01-toyama">富山県</label><br>
                <input id="radio01-ishiakwa" type="radio" name="radio01" value="金沢市">
                <label for="radio01-ishiakwa">石川県</label><br>
                <input id="radio01-fukui" type="radio" name="radio01" value="福井市">
                <label for="radio01-fukui">福井県</label>
            </div>
            <div>※IE11では下記に県庁所在地が表示されません。</div>
            <div id="result"></div>
            <div id="button-container">
                <button id="button01">表示</button>
            </div>
        </form>
    </div>
</main>