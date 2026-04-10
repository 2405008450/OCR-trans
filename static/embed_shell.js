(function () {
    const params = new URLSearchParams(window.location.search);
    const isEmbedded = params.get('embed') === '1' || window.self !== window.top;

    if (!isEmbedded) {
        return;
    }

    function markEmbedded() {
        document.documentElement.classList.add('embedded-shell');
        document.body.classList.add('embedded-shell');
    }

    function getElementHeight(element) {
        if (!element) {
            return 0;
        }

        const rect = element.getBoundingClientRect();
        return Math.max(
            element.scrollHeight || 0,
            element.offsetHeight || 0,
            Math.ceil(rect.height || 0)
        );
    }

    function measureHeight() {
        const primaryRoot =
            document.querySelector('.page, .container, main, .main-content') ||
            document.body.firstElementChild;
        const childHeights = Array.from(document.body.children).map((element) => getElementHeight(element));

        return Math.max(
            getElementHeight(primaryRoot),
            ...childHeights,
            0
        );
    }

    let lastNotifiedHeight = 0;

    function notifyParent() {
        if (window.parent === window) {
            return;
        }

        const nextHeight = measureHeight();
        if (lastNotifiedHeight > 0 && Math.abs(nextHeight - lastNotifiedHeight) <= 1) {
            return;
        }
        lastNotifiedHeight = nextHeight;

        window.parent.postMessage(
            {
                type: 'certificate-translation:resize',
                path: window.location.pathname,
                height: nextHeight,
            },
            window.location.origin
        );
    }

    function scheduleNotify() {
        window.requestAnimationFrame(() => {
            window.requestAnimationFrame(() => {
                notifyParent();
            });
        });
    }

    if (document.readyState === 'loading') {
        document.addEventListener(
            'DOMContentLoaded',
            () => {
                markEmbedded();
                scheduleNotify();
            },
            { once: true }
        );
    } else {
        markEmbedded();
        scheduleNotify();
    }

    window.addEventListener('load', scheduleNotify);
    window.addEventListener('resize', scheduleNotify);

    if (typeof ResizeObserver !== 'undefined') {
        const observer = new ResizeObserver(() => {
            scheduleNotify();
        });

        observer.observe(document.documentElement);
        observer.observe(document.body);
    }
})();
