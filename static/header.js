(function () {
    const APP_BUILD_ID = document.querySelector('meta[name="app-build-id"]')?.content || 'dev';
    const BUILD_STORAGE_KEY = 'app_build_id';
    const UPLOAD_TIMEOUT_MS = 240000;
    const PAGE_TRANSITION_DELAY_MS = 180;
    const REFRESH_HINT = '页面更新后若按钮异常、上传无响应或界面显示异常，请先按 Ctrl+F5 强制刷新；Mac 请按 Command+Shift+R。';
    const REFRESH_HINT_HTML = '页面更新后若按钮异常、上传无响应或界面显示异常，请先按 <kbd>Ctrl+F5</kbd> 强制刷新；Mac 请按 <kbd>Command+Shift+R</kbd>。';
    const RELEASE_NOTES = [
        '\u652f\u6301\u5bfc\u5165 .doc \u6587\u6863',
        '\u652f\u6301\u6279\u91cf\u4e0b\u8f7d\u5904\u7406\u5b8c\u6210\u6587\u4ef6',
        '4月20日_解除pdf2docx页数限制，文件大小限制为95M以下',
        '4月28日_增加通用证件翻译70+语种，记忆增加dsv4p模型'

    ];
    const originalFetch = window.fetch.bind(window);
    const navItems = [
        { href: '/', icon: 'fa-home', text: '首页' },
        { href: '/dashboard', icon: 'fa-gauge-high', text: '工作台' },
        { href: '/certificate-translation', icon: 'fa-id-card', text: '证件翻译聚合' },
        { href: '/pdf2docx', icon: 'fa-file-word', text: '不可编辑预处理' },
        { href: '/number-check', icon: 'fa-check-double', text: '数字专检' },
        { href: '/alignment', icon: 'fa-object-group', text: '多语对照' },
        { href: '/zhongfanyi', icon: 'fa-spell-check', text: '中翻专检' },
    ];
    const navActiveAliases = {
        '/certificate-translation': ['/certificate-translation', '/doc-translate', '/drivers-license', '/business-licence'],
    };
    let pageTransitionInProgress = false;

    const shell = window.AppShell || {};
    shell.buildId = APP_BUILD_ID;
    shell.refreshHint = REFRESH_HINT;
    shell.uploadTimeoutMs = UPLOAD_TIMEOUT_MS;
    shell.showToast = showToast;
    shell.showPageTransition = () => showPageTransition();
    shell.submitTaskRequest = (input, init = {}, options = {}) => fetchWithUploadTimeout(input, init, options);
    window.AppShell = shell;

    patchFetchForUploads();
    initPageTransitions();

    function patchFetchForUploads() {
        if (window.__APP_SHELL_UPLOAD_TIMEOUT_PATCHED__) {
            return;
        }

        window.__APP_SHELL_UPLOAD_TIMEOUT_PATCHED__ = true;
        window.fetch = function patchedFetch(input, init = {}) {
            if (!shouldApplyUploadTimeout(init)) {
                return originalFetch(input, init);
            }
            return fetchWithUploadTimeout(input, init, { nativeFetch: originalFetch });
        };
    }

    function shouldApplyUploadTimeout(init) {
        const method = String(init?.method || 'GET').toUpperCase();
        return method === 'POST' && typeof FormData !== 'undefined' && init?.body instanceof FormData;
    }

    async function fetchWithUploadTimeout(input, init = {}, options = {}) {
        const nativeFetch = options.nativeFetch || originalFetch;
        const timeoutMs = Number(options.timeoutMs) > 0 ? Number(options.timeoutMs) : UPLOAD_TIMEOUT_MS;
        const controller = new AbortController();
        const timeoutId = window.setTimeout(() => controller.abort('upload-timeout'), timeoutMs);

        if (init.signal) {
            if (init.signal.aborted) {
                controller.abort(init.signal.reason);
            } else {
                init.signal.addEventListener('abort', () => controller.abort(init.signal.reason), { once: true });
            }
        }

        try {
            return await nativeFetch(input, { ...init, signal: controller.signal });
        } catch (error) {
            throw normalizeUploadError(error, timeoutMs);
        } finally {
            window.clearTimeout(timeoutId);
        }
    }

    function normalizeUploadError(error, timeoutMs) {
        if (error?.name === 'AbortError') {
            return new Error(
                `上传等待已超过 ${Math.round(timeoutMs / 1000)} 秒，任务可能尚未创建。请检查网络后重试；若工作台没有新任务，说明请求还没有完整到达后端。`
            );
        }

        const message = String(error?.message || '');
        if (/Failed to fetch|Load failed|NetworkError/i.test(message)) {
            return new Error('上传链路已中断，文件可能还没有完整传到后端。请重试；若频繁出现，请检查服务器前的 Nginx 或负载均衡上传大小与超时配置。');
        }

        return error instanceof Error ? error : new Error(message || '请求失败');
    }

    function initPageTransitions() {
        const isEmbedded = window.self !== window.top || new URLSearchParams(window.location.search).get('embed') === '1';
        if (isEmbedded) {
            document.documentElement.classList.remove('app-page-preparing', 'app-page-leaving');
            return;
        }

        injectPageTransitionStyle();
        injectRuntimeShellOverrides();
        ensurePageTransitionOverlay();
        document.documentElement.classList.add('app-transition-enabled');

        window.addEventListener('pageshow', () => {
            markPageReady();
        });

        window.addEventListener('beforeunload', () => {
            showPageTransition();
        });

        document.addEventListener('click', handlePageLinkClick);
    }

    function scheduleMarkPageReady() {
        window.requestAnimationFrame(() => {
            window.requestAnimationFrame(markPageReady);
        });
    }

    function injectPageTransitionStyle() {
        if (document.getElementById('appTransitionStyle')) {
            return;
        }

        const style = document.createElement('style');
        style.id = 'appTransitionStyle';
        style.textContent = `
            html.app-transition-enabled {
                background: #040812;
            }
            html.app-page-preparing body {
                opacity: 0;
            }
            html.app-page-ready body {
                opacity: 1;
            }
            html.app-transition-enabled body > :not(.app-transition-overlay) {
                transition: opacity 0.18s ease;
            }
            html.app-page-leaving body {
                overflow: hidden;
            }
            html.app-page-leaving body > :not(.app-transition-overlay) {
                opacity: 0.96;
                pointer-events: none;
            }
            .app-transition-overlay {
                position: fixed;
                inset: 0;
                z-index: 10000;
                display: flex;
                align-items: center;
                justify-content: center;
                padding: 24px;
                opacity: 0;
                visibility: hidden;
                pointer-events: none;
                background:
                    radial-gradient(circle at 42% 36%, rgba(56, 189, 248, 0.18), transparent 28%),
                    rgba(3, 12, 24, 0.56);
                transition: opacity 0.16s ease, visibility 0.16s ease;
            }
            .app-transition-overlay.is-visible,
            html.app-page-leaving .app-transition-overlay {
                opacity: 1;
                visibility: visible;
                pointer-events: auto;
            }
            .app-transition-panel {
                display: grid;
                grid-template-columns: auto 1fr;
                align-items: center;
                gap: 16px;
                min-width: min(360px, calc(100vw - 48px));
                padding: 18px 20px;
                border-radius: 22px;
                border: 1px solid rgba(125, 211, 252, 0.26);
                background: linear-gradient(145deg, rgba(9, 22, 38, 0.88), rgba(15, 35, 58, 0.82));
                color: #f8fafc;
                box-shadow: 0 30px 80px rgba(2, 8, 23, 0.38);
            }
            .app-transition-spinner {
                width: 44px;
                height: 44px;
                border-radius: 999px;
                border: 4px solid rgba(186, 230, 253, 0.24);
                border-top-color: #7dd3fc;
                border-right-color: rgba(52, 211, 153, 0.86);
                animation: app-transition-spin 0.82s linear infinite;
                box-shadow: 0 0 30px rgba(56, 189, 248, 0.22);
            }
            .app-transition-copy {
                display: grid;
                gap: 4px;
                min-width: 0;
            }
            .app-transition-copy strong {
                font-size: 16px;
                line-height: 1.35;
            }
            .app-transition-copy span {
                color: rgba(226, 232, 240, 0.76);
                font-size: 13px;
                line-height: 1.5;
            }
            @keyframes app-transition-spin {
                to {
                    transform: rotate(360deg);
                }
            }
            @media (prefers-reduced-motion: reduce) {
                html.app-page-ready body {
                    animation: none;
                }
                html.app-transition-enabled body > :not(.app-transition-overlay),
                .app-transition-overlay {
                    transition: none;
                }
                .app-transition-spinner {
                    animation-duration: 1.4s;
                }
            }
        `;
        document.head.appendChild(style);
    }

    function injectRuntimeShellOverrides() {
        if (document.getElementById('appShellRuntimeOverrideStyle')) {
            return;
        }

        const style = document.createElement('style');
        style.id = 'appShellRuntimeOverrideStyle';
        style.textContent = `
            html.app-transition-enabled body,
            html.app-page-ready body,
            html.app-page-leaving body {
                animation: none !important;
                filter: none !important;
                transform: none !important;
            }
            html.app-transition-enabled body > :not(.app-transition-overlay) {
                transition: opacity 0.18s ease !important;
            }
            html.app-page-leaving body > :not(.app-transition-overlay) {
                filter: none !important;
                opacity: 0.96;
                transform: none !important;
            }
            .unified-top-nav a {
                line-height: 1;
                transform: none !important;
                transition: background-color 0.18s ease, border-color 0.18s ease, color 0.18s ease !important;
            }
            .unified-top-nav a:hover,
            .unified-top-nav a.active {
                transform: none !important;
            }
        `;
        document.head.appendChild(style);
    }

    function ensurePageTransitionOverlay() {
        let overlay = document.getElementById('appTransitionOverlay');
        if (overlay) {
            return overlay;
        }

        overlay = document.createElement('div');
        overlay.id = 'appTransitionOverlay';
        overlay.className = 'app-transition-overlay';
        overlay.setAttribute('aria-hidden', 'true');
        overlay.innerHTML = `
            <div class="app-transition-panel" role="status" aria-live="polite">
                <div class="app-transition-spinner" aria-hidden="true"></div>
                <div class="app-transition-copy">
                    <strong>正在切换页面</strong>
                    <span>请稍候，界面正在平滑加载</span>
                </div>
            </div>
        `;
        document.body.appendChild(overlay);
        return overlay;
    }

    function markPageReady() {
        const root = document.documentElement;
        pageTransitionInProgress = false;
        root.classList.remove('app-page-preparing', 'app-page-leaving');
        root.classList.add('app-page-ready');
        document.getElementById('appTransitionOverlay')?.classList.remove('is-visible');
    }

    function showPageTransition() {
        const root = document.documentElement;
        if (!root.classList.contains('app-transition-enabled')) {
            return;
        }
        pageTransitionInProgress = true;
        ensurePageTransitionOverlay().classList.add('is-visible');
        root.classList.remove('app-page-ready', 'app-page-preparing');
        root.classList.add('app-page-leaving');
    }

    function handlePageLinkClick(event) {
        const targetUrl = getTransitionTargetUrl(event);
        if (!targetUrl) {
            return;
        }

        event.preventDefault();
        if (pageTransitionInProgress) {
            return;
        }

        showPageTransition();
        window.setTimeout(() => {
            window.location.href = targetUrl.href;
        }, PAGE_TRANSITION_DELAY_MS);
    }

    function getTransitionTargetUrl(event) {
        if (event.defaultPrevented || event.button !== 0 || event.metaKey || event.ctrlKey || event.shiftKey || event.altKey) {
            return null;
        }

        const link = event.target.closest?.('a[href]');
        if (!link || link.dataset.noPageTransition === 'true' || link.closest('[data-no-page-transition]')) {
            return null;
        }

        const href = link.getAttribute('href') || '';
        const target = (link.getAttribute('target') || '').toLowerCase();
        if (
            !href ||
            href.startsWith('#') ||
            link.hasAttribute('download') ||
            (target && target !== '_self') ||
            /^(javascript|mailto|tel):/i.test(href)
        ) {
            return null;
        }

        let url;
        try {
            url = new URL(link.href, window.location.href);
        } catch (_) {
            return null;
        }

        if (url.origin !== window.location.origin || !/^https?:$/.test(url.protocol)) {
            return null;
        }

        const current = new URL(window.location.href);
        if (url.pathname === current.pathname && url.search === current.search) {
            return null;
        }

        return url;
    }

    function injectSharedStyle() {
        if (document.getElementById('appShellStyle')) {
            return;
        }

        const style = document.createElement('style');
        style.id = 'appShellStyle';
        style.textContent = `
            .unified-global-topbar {
                display: flex !important;
                justify-content: space-between !important;
                align-items: center !important;
                flex-wrap: wrap !important;
                gap: 16px !important;
                margin-bottom: 18px !important;
                padding-bottom: 18px !important;
                border-bottom: 1px solid rgba(255, 255, 255, 0.06) !important;
            }
            .unified-global-topbar .brand {
                display: inline-flex;
                align-items: center;
                gap: 14px;
            }
            .unified-top-nav {
                display: flex;
                flex-wrap: wrap;
                gap: 12px;
            }
            .unified-top-nav a {
                display: inline-flex;
                align-items: center;
                gap: 8px;
                min-height: 42px;
                line-height: 1;
                padding: 0 16px;
                border-radius: 999px;
                border: 1px solid rgba(255, 255, 255, 0.08);
                background: rgba(255, 255, 255, 0.04);
                text-decoration: none;
                color: #e2e8f0;
                font-size: 14px;
                font-weight: 500;
                transition: background-color 0.18s ease, border-color 0.18s ease, color 0.18s ease;
                backdrop-filter: blur(12px);
            }
            .unified-top-nav a:hover,
            .unified-top-nav a.active {
                background: rgba(56, 189, 248, 0.16);
                border-color: rgba(56, 189, 248, 0.4);
                color: #fff;
            }
            .page-hero-header {
                color: #fff;
                margin-bottom: 24px;
            }
            .shell-refresh-notice {
                display: grid;
                grid-template-columns: auto 1fr;
                align-items: flex-start;
                gap: 14px;
                margin-bottom: 20px;
                padding: 16px 18px;
                border-radius: 20px;
                border: 1px solid rgba(251, 191, 36, 0.42);
                background: linear-gradient(135deg, rgba(251, 191, 36, 0.24), rgba(249, 115, 22, 0.2));
                box-shadow: 0 20px 44px rgba(120, 53, 15, 0.2);
                position: relative;
                overflow: hidden;
            }
            .shell-refresh-notice::before {
                content: '';
                position: absolute;
                inset: 0 auto 0 0;
                width: 6px;
                background: linear-gradient(180deg, #facc15, #f97316);
            }
            .shell-refresh-notice .notice-icon {
                width: 42px;
                height: 42px;
                border-radius: 14px;
                display: grid;
                place-items: center;
                background: rgba(120, 53, 15, 0.18);
                color: #fff7ed;
                box-shadow: inset 0 0 0 1px rgba(255, 247, 237, 0.12);
            }
            .shell-refresh-notice .notice-copy {
                min-width: 0;
            }
            .shell-refresh-notice .notice-badge {
                display: inline-flex;
                align-items: center;
                min-height: 24px;
                padding: 0 10px;
                margin-bottom: 8px;
                border-radius: 999px;
                background: rgba(120, 53, 15, 0.22);
                color: #fff7ed;
                font-size: 12px;
                font-weight: 800;
                letter-spacing: 0.08em;
            }
            .shell-refresh-notice .notice-title {
                display: block;
                margin-bottom: 6px;
                color: #fff7ed;
                font-size: 17px;
                font-weight: 800;
                line-height: 1.35;
            }
            .shell-refresh-notice .notice-text {
                color: rgba(255, 247, 237, 0.96);
                font-size: 14px;
                font-weight: 600;
                line-height: 1.7;
            }
            .shell-refresh-notice kbd {
                display: inline-flex;
                align-items: center;
                min-height: 26px;
                margin: 0 3px;
                padding: 0 9px;
                border-radius: 8px;
                border: 1px solid rgba(255, 247, 237, 0.32);
                background: rgba(120, 53, 15, 0.3);
                color: #ffffff;
                font-size: 12px;
                font-weight: 800;
                font-family: Consolas, "SFMono-Regular", "Liberation Mono", Menlo, monospace;
                box-shadow: inset 0 -1px 0 rgba(255, 247, 237, 0.16);
                vertical-align: middle;
            }
            .shell-release-note {
                display: flex;
                align-items: center;
                flex-wrap: wrap;
                gap: 14px;
                margin-bottom: 18px;
                padding: 14px 18px;
                border-radius: 14px;
                border: 1px solid rgba(14, 165, 233, 0.18);
                background: linear-gradient(135deg, rgba(11, 31, 52, 0.6), rgba(6, 17, 29, 0.5));
                box-shadow: 0 12px 28px rgba(0, 0, 0, 0.2);
                position: relative;
                overflow: hidden;
            }
            .shell-release-note::before {
                content: '';
                position: absolute;
                inset: 0 0 auto 0;
                height: 4px;
                background: linear-gradient(90deg, #0ea5e9, #38bdf8);
            }
            .shell-release-note .release-icon {
                width: 40px;
                height: 40px;
                border-radius: 12px;
                display: grid;
                place-items: center;
                background: linear-gradient(135deg, #0284c7, #0369a1);
                color: #ffffff;
                font-size: 18px;
                flex: 0 0 auto;
            }
            .shell-release-note .release-copy {
                min-width: 0;
                flex: 1 1 0;
                display: flex;
                align-items: center;
                flex-wrap: wrap;
                gap: 10px 12px;
            }
            .shell-release-note .release-badge {
                display: inline-flex;
                align-items: center;
                min-height: 24px;
                padding: 0 10px;
                margin: 0;
                border-radius: 999px;
                background: rgba(14, 165, 233, 0.18) !important;
                border: 1px solid rgba(14, 165, 233, 0.32) !important;
                color: #e0f2fe !important;
                -webkit-text-fill-color: currentColor !important;
                font-size: 12px;
                font-weight: 900;
                letter-spacing: 0.06em;
                opacity: 1 !important;
                flex: 0 0 auto;
            }
            .shell-release-note .release-title {
                display: inline;
                margin: 0;
                color: #f0f9ff !important;
                -webkit-text-fill-color: currentColor !important;
                font-size: 15px;
                font-weight: 800;
                line-height: 1.5;
                text-shadow: none;
                opacity: 1 !important;
                flex: 0 1 auto;
            }
            .shell-release-note .release-list {
                display: flex;
                align-items: center;
                flex-wrap: wrap;
                gap: 8px;
                margin: 0;
                padding: 0;
                list-style: none;
                flex: 1 1 100%;
            }
            .shell-release-note .release-list li {
                display: inline-flex;
                align-items: center;
                gap: 8px;
                padding: 6px 10px;
                border-radius: 999px;
                background: rgba(56, 189, 248, 0.12) !important;
                border: 1px solid rgba(56, 189, 248, 0.16) !important;
                color: #bae6fd !important;
                -webkit-text-fill-color: currentColor !important;
                opacity: 1 !important;
                font-size: 14px;
                font-weight: 700;
                line-height: 1.4;
                text-shadow: none;
            }
            .shell-release-note .release-list li::before {
                content: '';
                width: 6px;
                height: 6px;
                border-radius: 999px;
                background: #38bdf8;
                flex: 0 0 auto;
            }
            .shell-build-toast {
                position: fixed;
                right: 18px;
                bottom: 18px;
                z-index: 9999;
                display: flex;
                align-items: center;
                gap: 12px;
                max-width: min(460px, calc(100vw - 28px));
                padding: 14px 16px;
                border-radius: 16px;
                border: 1px solid rgba(125, 211, 252, 0.26);
                background: rgba(15, 23, 42, 0.94);
                box-shadow: 0 24px 50px rgba(2, 8, 23, 0.32);
                color: #f8fafc;
                backdrop-filter: blur(14px);
            }
            .shell-build-toast button {
                border: none;
                background: rgba(255, 255, 255, 0.08);
                color: #e2e8f0;
                width: 32px;
                height: 32px;
                border-radius: 999px;
                cursor: pointer;
            }
            .shell-build-toast button:hover {
                background: rgba(255, 255, 255, 0.16);
            }
            select option {
                background-color: #0d2138;
                color: #f6f8fb;
            }
            @media (max-width: 720px) {
                .shell-refresh-notice {
                    border-radius: 16px;
                    padding: 14px 14px 14px 16px;
                    gap: 12px;
                }
                .shell-refresh-notice .notice-title {
                    font-size: 15px;
                }
                .shell-refresh-notice .notice-text {
                    font-size: 13px;
                }
                .shell-release-note {
                    align-items: flex-start;
                    border-radius: 12px;
                    padding: 14px 14px 14px 16px;
                    gap: 12px;
                }
                .shell-release-note .release-copy {
                    display: grid;
                    gap: 8px;
                }
                .shell-release-note .release-title {
                    font-size: 14px;
                }
                .shell-release-note .release-list {
                    gap: 7px;
                }
                .shell-release-note .release-list li {
                    font-size: 13px;
                    padding: 6px 9px;
                }
                .shell-build-toast {
                    right: 10px;
                    bottom: 10px;
                    left: 10px;
                    max-width: none;
                }
            }
        `;
        document.head.appendChild(style);
    }

    function renderNav() {
        injectSharedStyle();

        const isEmbedded = window.self !== window.top || new URLSearchParams(window.location.search).get('embed') === '1';
        if (isEmbedded) {
            document.querySelectorAll('header.header, header.topbar').forEach((header) => {
                header.style.display = 'none';
            });
            return;
        }

        const shellSlot = document.getElementById('appShellSlot');
        const shellHost = shellSlot?.querySelector('.app-shell-inner') || shellSlot || document.querySelector('.container, .page');
        if (!shellHost) {
            scheduleMarkPageReady();
            return;
        }

        const currentPath = window.location.pathname;
        const navHtml = navItems
            .map((item) => {
                const activePrefixes = navActiveAliases[item.href] || [item.href];
                const isActive = activePrefixes.some((prefix) => currentPath === prefix || (prefix !== '/' && currentPath.startsWith(prefix + '/')));
                return `<a href="${item.href}" class="${isActive ? 'active' : ''}"><i class="fas ${item.icon}"></i> ${item.text}</a>`;
            })
            .join('');

        if (!shellHost.querySelector('.unified-global-topbar')) {
            const unifiedTopbar = document.createElement('header');
            unifiedTopbar.className = 'topbar unified-global-topbar';
            unifiedTopbar.innerHTML = `
                <div class="brand">
                    <div class="brand-mark" style="width: 44px; height: 44px; border-radius: 14px; display: grid; place-items: center; background: linear-gradient(135deg, rgba(56, 189, 248, 0.28), rgba(56, 189, 248, 0.62)); font-size: 20px; color: #fff;">
                        <i class="fas fa-layer-group"></i>
                    </div>
                    <div style="font-weight: 700; font-size: 17px; color: #fff;">文档处理工作台</div>
                </div>
                <nav class="unified-top-nav">
                    ${navHtml}
                </nav>
            `;
            shellHost.insertBefore(unifiedTopbar, shellHost.firstChild);
        }

        document.querySelectorAll('header.topbar:not(.unified-global-topbar)').forEach((header) => header.remove());
        document.querySelectorAll('header.header').forEach((header) => {
            header.querySelector('nav')?.remove();
            header.classList.add('page-hero-header');
        });

        injectRefreshNotice(shellHost);
        injectReleaseNotice(shellHost);
        announceBuildUpdate();
        scheduleMarkPageReady();
    }

    function injectRefreshNotice(topContainer) {
        if (topContainer.querySelector('.shell-refresh-notice')) {
            return;
        }

        const notice = document.createElement('section');
        notice.className = 'shell-refresh-notice';
        notice.innerHTML = `
            <div class="notice-icon"><i class="fas fa-triangle-exclamation"></i></div>
            <div class="notice-copy">
                <div class="notice-badge">使用提示</div>
                <strong class="notice-title">页面异常时先强制刷新缓存</strong>
                <div class="notice-text">${REFRESH_HINT_HTML}</div>
            </div>
        `;

        const topbar = topContainer.querySelector('.unified-global-topbar');
        if (topbar?.nextSibling) {
            topContainer.insertBefore(notice, topbar.nextSibling);
        } else {
            topContainer.appendChild(notice);
        }
    }

    function injectReleaseNotice(topContainer) {
        if (topContainer.querySelector('.shell-release-note')) {
            return;
        }

        const note = document.createElement('section');
        note.className = 'shell-release-note';
        note.innerHTML = `
            <div class="release-icon"><i class="fas fa-bullhorn"></i></div>
            <div class="release-copy">
                <div class="release-badge">\u6700\u8fd1\u66f4\u65b0</div>
                <strong class="release-title">\u5f53\u524d\u7248\u672c\u5df2\u540c\u6b65\u4ee5\u4e0b\u5185\u5bb9</strong>
                <ul class="release-list">
                    ${RELEASE_NOTES.map((item) => `<li>${item}</li>`).join('')}
                </ul>
            </div>
        `;

        const anchor = topContainer.querySelector('.shell-refresh-notice') || topContainer.querySelector('.unified-global-topbar');
        if (anchor?.nextSibling) {
            topContainer.insertBefore(note, anchor.nextSibling);
        } else {
            topContainer.appendChild(note);
        }
    }

    function announceBuildUpdate() {
        let previousBuildId = null;
        try {
            previousBuildId = window.localStorage.getItem(BUILD_STORAGE_KEY);
            window.localStorage.setItem(BUILD_STORAGE_KEY, APP_BUILD_ID);
        } catch (_) {
            return;
        }

        if (previousBuildId && previousBuildId !== APP_BUILD_ID) {
            showToast(`系统已更新。${REFRESH_HINT}`);
        }
    }

    function showToast(message) {
        document.querySelector('.shell-build-toast')?.remove();

        const toast = document.createElement('div');
        toast.className = 'shell-build-toast';
        toast.innerHTML = `
            <div style="flex:1;line-height:1.6;font-size:13px;">${message}</div>
            <button type="button" aria-label="关闭提示"><i class="fas fa-xmark"></i></button>
        `;

        const close = () => toast.remove();
        toast.querySelector('button')?.addEventListener('click', close);
        document.body.appendChild(toast);
        window.setTimeout(close, 12000);
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', renderNav);
    } else {
        renderNav();
    }
})();
