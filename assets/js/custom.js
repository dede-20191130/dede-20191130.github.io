window.addEventListener("DOMContentLoaded", (event) => {
    // 切替可能GIFを含むimg要素の設定
    const images = document.querySelectorAll('img.togglable-gif-image');
    images.forEach((image) => {
        image.addEventListener('click', toggleGifImage);
    });
    // ページ内アンカージャンプ時
    // ヘッダ部分によるスクロールのずれのための補正
    const pageAnchors = document.querySelectorAll('a[href^="#"]');
    pageAnchors.forEach((anchor) => {
        anchor.addEventListener('click', correntScroll);
    });

})


// gifアニメーションへの切り替え
function toggleGifImage() {
    const image = this;
    const src = image.src;
    const before = image.getAttribute("data-before");
    image.setAttribute('data-before', src);

    image.src = before ? before : src.substr(0, src.lastIndexOf(".")) + ".gif";
}

// ページ内アンカージャンプ時
// ヘッダ部分によるスクロールのずれのための補正
// ref:https://senoweb.jp/note/fixheader-anchorlink/
function correntScroll(event) {
    // 対象アンカー
    const href = event.currentTarget.getAttribute("href");
    if (href === "#" || href === "") {
        return;
    }
    // ヘッダ高さ
    const hdH = document.querySelector("header")?.clientHeight || 0;
    // ジャンプ先要素のdocumentに対する位置（Y）の取得
    const target = document.querySelector(href);
    const positionY = window.pageYOffset + target.getBoundingClientRect().top - hdH;
    // スクロールを指定
    window.scrollTo({
        top: positionY,
    })

    event.preventDefault();

}
// parentURLがバグっぽいのでテーマを上書き
function loadSvg(file, parent, path = iconsPath) {
    const link = `{{ absURL "" }}${path}${file}.svg`;
    fetch(link)
        .then((response) => {
            return response.text();
        })
        .then((data) => {
            parent.innerHTML = data;
        });
}