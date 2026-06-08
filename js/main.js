/* ===== 数据管理 ===== */

// 用 category 标记：material / gallery / video
const POSTS = [
    {
        id: 1,
        title: "ETF 份额统计数据",
        summary: "最新的 ETF 份额统计表格，可直接下载查看。",
        tag: "数据",
        category: "material",
        icon: "📊",
        date: "2026-06-05",
        link: "materials.html"
    },
    {
        id: 2,
        title: "编程学习路线图",
        summary: "从入门到进阶的系统学习路径，涵盖前端、后端和全栈。",
        tag: "笔记",
        category: "material",
        icon: "📘",
        date: "2026-06-08",
        link: "materials.html"
    },
    {
        id: 3,
        title: "设计资源合集",
        summary: "收集的各种设计素材、参考网站和配色工具。",
        tag: "资源",
        category: "material",
        icon: "🎨",
        date: "2026-06-07",
        link: "materials.html"
    },
    {
        id: 4,
        title: "旅行摄影精选",
        summary: "旅途中记录下的风景和人文瞬间。",
        tag: "摄影",
        category: "gallery",
        icon: "📷",
        date: "2026-06-06",
        link: "gallery.html"
    }
];

const GALLERY = [
    // 示例： { src: "assets/pics/photo1.jpg", caption: "描述文字" }
];

const VIDEOS = [
    // 示例： { src: "assets/videos/video1.mp4", title: "标题", desc: "描述" }
    // 也可以用外部链接： { src: "https://www.youtube.com/embed/xxx", title: "标题", desc: "描述", external: true }
];

const MATERIALS = [
    { icon: "📊", title: "ETF 份额统计", desc: "最新的 ETF 份额统计数据，包含各基金份额变化情况。", tags: ["ETF", "数据", "Excel"], date: "2026-06-05", file: "assets/posts/ETF份额统计.xlsx", fileName: "ETF份额统计.xlsx" },
    // 示例： { icon: "📄", title: "标题", desc: "描述", tags: ["标签1"], date: "2026-06-08", file: "文件路径", fileName: "文档名.xlsx" }
];

/* ===== 渲染函数 ===== */

function renderRecentPosts() {
    const container = document.getElementById("recent-posts");
    if (!container) return;

    const items = POSTS.slice(0, 6);
    if (items.length === 0) {
        container.innerHTML = `<div class="empty-state">
            <div class="emoji">📝</div>
            <h3>还没有内容</h3>
            <p>内容会通过 JavaScript 动态加载，快去添加吧！</p>
        </div>`;
        return;
    }

    container.innerHTML = items.map(post => {
        const imgHtml = post.cover
            ? `<img class="post-card-image" src="${post.cover}" alt="${post.title}" loading="lazy">`
            : `<div class="post-card-image" style="display:flex;align-items:center;justify-content:center;font-size:2.5rem;color:var(--text-secondary);background:var(--bg-alt)">${post.icon || '📄'}</div>`;

        return `<a class="post-card" href="${post.link}">
            ${imgHtml}
            <div class="post-card-body">
                <span class="post-card-tag">${post.tag || post.category}</span>
                <h3>${post.title}</h3>
                <p>${post.summary}</p>
                <div class="post-card-date">${post.date}</div>
            </div>
        </a>`;
    }).join("");
}

function renderGallery() {
    const container = document.getElementById("gallery-grid");
    if (!container) return;

    if (GALLERY.length === 0) {
        container.innerHTML = `<div class="empty-state">
            <div class="emoji">🖼️</div>
            <h3>图库还是空的</h3>
            <p>在 <code>js/main.js</code> 的 <code>GALLERY</code> 数组里添加图片路径吧！</p>
        </div>`;
        return;
    }

    container.innerHTML = GALLERY.map((item, i) =>
        `<div class="gallery-item" onclick="openLightbox(${i})">
            <img src="${item.src}" alt="${item.caption}" loading="lazy">
            <div class="caption">${item.caption}</div>
        </div>`
    ).join("");
}

function renderVideos() {
    const container = document.getElementById("video-grid");
    if (!container) return;

    if (VIDEOS.length === 0) {
        container.innerHTML = `<div class="empty-state">
            <div class="emoji">🎬</div>
            <h3>还没有视频</h3>
            <p>在 <code>js/main.js</code> 的 <code>VIDEOS</code> 数组里添加视频吧！</p>
        </div>`;
        return;
    }

    container.innerHTML = VIDEOS.map(v => {
        const player = v.external
            ? `<iframe src="${v.src}" frameborder="0" allowfullscreen loading="lazy"></iframe>`
            : `<video src="${v.src}" controls preload="metadata"></video>`;
        return `<div class="video-card">
            <div class="video-wrapper">${player}</div>
            <div class="video-card-body">
                <h3>${v.title}</h3>
                ${v.desc ? `<p>${v.desc}</p>` : ''}
            </div>
        </div>`;
    }).join("");
}

function renderMaterials() {
    const container = document.getElementById("materials-list");
    if (!container) return;

    if (MATERIALS.length === 0) {
        container.innerHTML = `<div class="empty-state">
            <div class="emoji">📚</div>
            <h3>还没有资料</h3>
            <p>在 <code>js/main.js</code> 的 <code>MATERIALS</code> 数组里添加内容吧！</p>
        </div>`;
        return;
    }

    container.innerHTML = MATERIALS.map(m => {
        const hasFile = m.file && m.fileName;
        const previewUrl = hasFile ? `preview.html?file=${encodeURIComponent(m.file)}&name=${encodeURIComponent(m.fileName)}` : '';
        const downloadUrl = hasFile ? m.file : '';

        let inner = `
            <div class="material-icon">${m.icon || '📄'}</div>
            <div class="material-content">
                <h3>${m.title}</h3>
                <p>${m.desc}</p>
                <div class="material-meta">
                    <span>${m.date}</span>
                    ${(m.tags || []).map(t => `<span class="material-tag">${t}</span>`).join('')}
                </div>
            </div>
        `;

        if (hasFile) {
            inner += `
                <div class="material-actions">
                    <a href="${previewUrl}" class="mat-action mat-preview">👁️ 预览</a>
                    <a href="${downloadUrl}" class="mat-action mat-dl" download onclick="event.stopPropagation()">⬇️</a>
                </div>
            `;
        }

        if (previewUrl) {
            return `<a class="material-item" href="${previewUrl}">${inner}</a>`;
        }
        return `<div class="material-item">${inner}</div>`;
    }).join("");
}

/* ===== Lightbox ===== */

function openLightbox(index) {
    const item = GALLERY[index];
    if (!item) return;

    const lb = document.getElementById("lightbox") || createLightbox();
    lb.innerHTML = `
        <span class="lightbox-close" onclick="closeLightbox()">&times;</span>
        <img src="${item.src}" alt="${item.caption}">
    `;
    lb.classList.add("active");
    document.body.style.overflow = "hidden";
}

function closeLightbox() {
    const lb = document.getElementById("lightbox");
    if (lb) {
        lb.classList.remove("active");
        document.body.style.overflow = "";
    }
}

function createLightbox() {
    const lb = document.createElement("div");
    lb.id = "lightbox";
    lb.className = "lightbox";
    lb.addEventListener("click", closeLightbox);
    document.body.appendChild(lb);
    return lb;
}

// ESC 关闭 lightbox
document.addEventListener("keydown", e => {
    if (e.key === "Escape") closeLightbox();
});

/* ===== 移动端菜单 ===== */

document.addEventListener("DOMContentLoaded", () => {
    const toggle = document.querySelector(".menu-toggle");
    const nav = document.querySelector(".nav-links");
    if (toggle && nav) {
        toggle.addEventListener("click", () => nav.classList.toggle("open"));
        // 点击链接关闭菜单
        nav.querySelectorAll("a").forEach(a => {
            a.addEventListener("click", () => nav.classList.remove("open"));
        });
    }

    // 自动检测当前页面高亮导航
    const currentPath = window.location.pathname.split('/').pop() || 'index.html';
    document.querySelectorAll(".nav-links a").forEach(a => {
        const href = a.getAttribute("href");
        if (href === currentPath || (currentPath === "" && href === "/")) {
            a.classList.add("active");
        } else {
            a.classList.remove("active");
        }
    });

    // 渲染各页面
    renderRecentPosts();
    renderGallery();
    renderVideos();
    renderMaterials();
});
