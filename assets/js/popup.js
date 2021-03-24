function popupImage() {
    const popup = document.getElementById('js-popup');
    if (!popup) return;

    const blackBg = document.getElementById('js-black-bg');
    const closeBtn = document.getElementById('js-close-btn');
    const showElems = document.getElementsByClassName('js-show-popup');

    closePopUp(blackBg);
    closePopUp(closeBtn);
    for (const e of showElems) {
        createPopUp(e);
    }
    
    function closePopUp(elem) {
        if (!elem) return;
        elem.addEventListener('click', () => {
            // // モーダルが閉じられたら、スライドを再開
            popup.classList.toggle('is-show');
            slideshow.resume();
        });
    }
    function createPopUp(elem) {
        if (!elem) return;
        const imgElems = elem.getElementsByTagName("img")
        for (const imgElem of imgElems) {
            imgElem.addEventListener('click', () => {
                const src = imgElem.src;
                if (!src) return;
                // // モーダルが開いたら、スライドを停止
                popedImg.src = src;
                popup.classList.toggle('is-show');
                slideshow.pause();
            });

        }
    }

}
const popedImg = document.getElementById('js-popup-img');
popupImage();