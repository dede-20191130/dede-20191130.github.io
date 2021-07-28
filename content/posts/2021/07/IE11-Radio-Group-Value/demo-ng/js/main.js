// not working in IE11
window.addEventListener("DOMContentLoaded", function () {
    document.getElementById("button01").onclick = function () {
        const radio01Value = document.forms["my-form"].radio01.value;
        if (!radio01Value) return;
        document.getElementById("result").innerHTML =
            "県庁所在地：" + radio01Value;
    };
});