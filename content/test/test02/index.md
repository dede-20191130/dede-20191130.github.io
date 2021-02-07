---
title: "Test02"
author: dede-20191130
date: 2020-11-23T14:45:27+09:00
# slug: test02
draft: true
toc: true
# type: test
animal: cat
cost:
    costAnhour: 120
lastmod: 2020-11-05
---

<style>
    input {
        appearance: none;
        -webkit-appearance: none;
        -moz-appearance: none;
        -o-appearance: none;
        border: none;
        outline: 1px solid #e8e4da;
        margin: 0;
        width: 15px;
        height: 15px;
        line-height: 15px;
        font-size: 10px;
        text-align: center;
        vertical-align: top;
        background-color: #ffaeae;
    }
    input[type="checkbox"]:checked::before {
        content: "済";
    }
}
</style>

## test list page
[https://link](../)

## H2見出し {#h2_headline .hoge style="color:red;" }

<!-- dtリストテスト
:  foo

gggooo
:  bar

脚注テスト-Hugo[^1]

[^1]: 最速の静的サイトジェネレーター

aiu[^2]
[^2]: あいういうおい


www.example.com

<p><del>やあ</del> こんにちは</p>

- [ ] hoge
- [x] fuga

<ul>
    <li><input disabled="" type="checkbox">hoge</li>
    <li><input checked="" disabled="" type="checkbox">fuga</li>
</ul>

<div >
    <ul style="appearance:auto; color:blue;">
        <li><input  type="checkbox">hoge</li>
        <li><input checked=""  type="checkbox" style="appearance:auto">fuga</li>
    </ul>

</div> -->

{{< param Animal >}}<br>

{{< param cost >}}  <br>

{{< param cost.costAnhour >}}  <br>

{{< param "cost.costAnhour" >}}  <br>