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

    function measureHeight() {
        return Math.max(
            document.documentElement.scrollHeight,
            document.documentElement.offsetHeight,
            document.body.scrollHeight,
            document.body.offsetHeight
        );
    }

    function notifyParent() {
        if (window.parent === window) {
            return;
        }

        window.parent.postMessage(
            {
                type: 'certificate-translation:resize',
                path: window.location.pathname,
                height: measureHeight(),
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
