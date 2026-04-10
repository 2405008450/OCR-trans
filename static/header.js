(function () {
    const APP_BUILD_ID = document.querySelector('meta[name="app-build-id"]')?.content || 'dev';
    const BUILD_STORAGE_KEY = 'app_build_id';
    const UPLOAD_TIMEOUT_MS = 240000;
    const REFRESH_HINT = '页面更新后若按钮异常、上传无响应或界面显示异常，请先按 Ctrl+F5 强制刷新；Mac 请按 Command+Shift+R。';
    const originalFetch = window.fetch.bind(window);
    const navItems = [
        { href: '/', icon: 'fa-home', text: '首页' },
        { href: '/dashboard', icon: 'fa-gauge-high', text: '工作台' },
        { href: '/certificate-translation', icon: 'fa-id-card', text: '证件翻译聚合' },
        { href: '/pdf2docx', icon: 'fa-file-word', text: '文档预处理' },
        { href: '/number-check', icon: 'fa-check-double', text: '数字专检' },
        { href: '/alignment', icon: 'fa-object-group', text: '多语对照' },
        { href: '/zhongfanyi', icon: 'fa-spell-check', text: '中翻专检' },
    ];

    const shell = window.AppShell || {};
    shell.buildId = APP_BUILD_ID;
    shell.refreshHint = REFRESH_HINT;
    shell.uploadTimeoutMs = UPLOAD_TIMEOUT_MS;
    shell.showToast = showToast;
    shell.submitTaskRequest = (input, init = {}, options = {}) => fetchWithUploadTimeout(input, init, options);
    window.AppShell = shell;

    patchFetchForUploads();

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
                padding: 0 16px;
                border-radius: 999px;
                border: 1px solid rgba(255, 255, 255, 0.08);
                background: rgba(255, 255, 255, 0.04);
                text-decoration: none;
                color: #e2e8f0;
                font-size: 14px;
                font-weight: 500;
                transition: all 0.24s ease;
                backdrop-filter: blur(12px);
            }
            .unified-top-nav a:hover,
            .unified-top-nav a.active {
                background: rgba(56, 189, 248, 0.16);
                border-color: rgba(56, 189, 248, 0.4);
                color: #fff;
                transform: translateY(-1.5px);
            }
            .page-hero-header {
                color: #fff;
                margin-bottom: 24px;
            }
            .shell-refresh-notice {
                display: flex;
                align-items: flex-start;
                gap: 12px;
                margin-bottom: 20px;
                padding: 14px 16px;
                border-radius: 18px;
                border: 1px solid rgba(125, 211, 252, 0.22);
                background: linear-gradient(135deg, rgba(14, 165, 233, 0.14), rgba(15, 23, 42, 0.35));
                box-shadow: 0 18px 40px rgba(2, 8, 23, 0.18);
            }
            .shell-refresh-notice i {
                margin-top: 2px;
                color: #7dd3fc;
            }
            .shell-refresh-notice strong {
                display: block;
                margin-bottom: 4px;
                color: #f8fafc;
                font-size: 14px;
            }
            .shell-refresh-notice span {
                color: rgba(226, 232, 240, 0.84);
                font-size: 13px;
                line-height: 1.7;
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

        const topContainer = document.querySelector('.container, .page');
        if (!topContainer) {
            return;
        }

        const currentPath = window.location.pathname;
        const navHtml = navItems
            .map((item) => {
                const isActive = currentPath === item.href || currentPath.startsWith(item.href + '/');
                return `<a href="${item.href}" class="${isActive ? 'active' : ''}"><i class="fas ${item.icon}"></i> ${item.text}</a>`;
            })
            .join('');

        if (!topContainer.querySelector('.unified-global-topbar')) {
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
            topContainer.insertBefore(unifiedTopbar, topContainer.firstChild);
        }

        document.querySelectorAll('header.topbar:not(.unified-global-topbar)').forEach((header) => header.remove());
        document.querySelectorAll('header.header').forEach((header) => {
            header.querySelector('nav')?.remove();
            header.classList.add('page-hero-header');
        });

        injectRefreshNotice(topContainer);
        announceBuildUpdate();
    }

    function injectRefreshNotice(topContainer) {
        if (topContainer.querySelector('.shell-refresh-notice')) {
            return;
        }

        const notice = document.createElement('section');
        notice.className = 'shell-refresh-notice';
        notice.innerHTML = `
            <i class="fas fa-rotate"></i>
            <div>
                <strong>使用提示</strong>
                <span>${REFRESH_HINT}</span>
            </div>
        `;

        const topbar = topContainer.querySelector('.unified-global-topbar');
        if (topbar?.nextSibling) {
            topContainer.insertBefore(notice, topbar.nextSibling);
        } else {
            topContainer.appendChild(notice);
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
            showToast('系统已更新。如果页面仍旧异常，请先按 Ctrl+F5 强制刷新。');
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
