document.addEventListener('DOMContentLoaded', function () {
    const image_view = document.querySelector('.image-view');
    const image_view_image = document.querySelector('.image-view-image');
    const close_btn = document.querySelector('.close-btn');
    const large_view = document.querySelectorAll('.large-view');

    if (image_view && image_view_image && close_btn && large_view) {
        close_btn.addEventListener('click', function () {
            image_view.classList.add('hidden');
        });

        for (var i = 0; i < large_view.length; i++) {
            large_view[i].addEventListener('click', function () {
                image_view_image.src = this.src;
                image_view.classList.remove('hidden');
                close_btn.classList.remove('hidden');
            });
        }
    }
});