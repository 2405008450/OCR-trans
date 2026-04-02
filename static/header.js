// 通用极致一体化导航栏及全局样式注入
(function() {
    const navItems = [
        { href: '/', icon: 'fa-home', text: '首页' },
        { href: '/dashboard', icon: 'fa-gauge-high', text: '工作台' },
        { href: '/certificate-translation', icon: 'fa-id-card', text: '证件翻译聚合' },
        { href: '/pdf2docx', icon: 'fa-file-word', text: '文档预处理' },
        { href: '/number-check', icon: 'fa-check-double', text: '数字专检' },
        { href: '/alignment', icon: 'fa-object-group', text: '多语对照' },
        { href: '/zhongfanyi', icon: 'fa-spell-check', text: '中翻专检' }
    ];

    function renderNav() {
        const isEmbedded = window.self !== window.top || new URLSearchParams(window.location.search).get('embed') === '1';

        if (isEmbedded) {
            // 在 iframe 中不渲染全局顶栏，且需移除原有的所有 Title Header，使嵌页做到纯净无死角
            const oldHeaders = document.querySelectorAll('header.header, header.topbar');
            oldHeaders.forEach(h => h.style.display = 'none');
            
            // 依然需要修正下拉框色盲 Bug
            const style = document.createElement('style');
            style.innerHTML = `
                select option {
                    background-color: #0d2138;
                    color: #f6f8fb;
                }
            `;
            document.head.appendChild(style);
            return;
        }

        const currentPath = window.location.pathname;
        let navHtml = '';

        navItems.forEach(item => {
            const isActive = currentPath === item.href || currentPath.startsWith(item.href + '/');
            navHtml += `<a href="${item.href}" class="${isActive ? 'active' : ''}"><i class="fas ${item.icon}"></i> ${item.text}</a>`;
        });

        const topContainer = document.querySelector('.container, .page');
        if (!topContainer) return;

        // 主动构建真正的全局绝对顶栏
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

        // 强行置于容器顶部
        topContainer.insertBefore(unifiedTopbar, topContainer.firstChild);

        // 切除杂乱的旧有头部逻辑
        const oldHeaders = document.querySelectorAll('header.header, header.topbar:not(.unified-global-topbar)');
        oldHeaders.forEach(header => {
            if (header.classList.contains('topbar')) {
                // 原有的独立 topbar（如 dashboard）直接拔除
                header.remove();
            } else if (header.classList.contains('header')) {
                // 原有的图文头（如 pdf2docx），剔除右侧 nav，保留它作为本页特有的大字报
                const oldNav = header.querySelector('nav');
                if (oldNav) oldNav.remove();
                header.className = 'page-hero-header'; // 改名降级
            }
        });

        // 注入包含 Navbar 和降级 Header 的公共强一致 CSS
        const style = document.createElement('style');
        style.innerHTML = `
            .unified-global-topbar {
                display: flex !important;
                justify-content: space-between !important;
                align-items: center !important;
                flex-wrap: wrap !important;
                gap: 16px !important;
                margin-bottom: 28px !important;
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
            /* 保护原有的降级 Banner 图文展示位 */
            .page-hero-header {
                color: #fff;
                margin-bottom: 24px;
            }
            /* 修复各种暗度下拉菜单的色盲现象 */
            select option {
                background-color: #0d2138;
                color: #f6f8fb;
            }
        `;
        document.head.appendChild(style);
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', renderNav);
    } else {
        renderNav();
    }
})();
