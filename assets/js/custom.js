// 切替可能GIFを含むimg要素の設定
window.addEventListener("DOMContentLoaded", (event) => {
    const images = document.querySelectorAll('img.togglable-gif-image');
    images.forEach((image) => {
        image.addEventListener('click', toggleGifImage);
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