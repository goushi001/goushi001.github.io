# 🍎 阿奇的资料站

个人博客网站，支持展示资料、图片和视频。

## 快速启动

```bash
cd blog
python3 -m http.server 3000
# 浏览器打开 http://localhost:3000
```

## 添加内容

所有内容通过 `js/main.js` 中的数组管理：

### 📄 资料
在 `MATERIALS` 数组里添加条目：
```js
{ icon: "📘", title: "标题", desc: "描述文字", tags: ["标签1", "标签2"], date: "2026-06-08" }
```

### 🖼️ 图库
图片丢进 `assets/pics/`，在 `GALLERY` 数组里引用：
```js
{ src: "assets/pics/photo1.jpg", caption: "描述文字" }
```
点击图片会自动弹出 Lightbox 大图查看。

### 🎬 视频
视频丢进 `assets/videos/`，在 `VIDEOS` 数组里引用：
```js
// 本地视频
{ src: "assets/videos/video1.mp4", title: "标题", desc: "描述" }
// 外部视频 (YouTube/B站)
{ src: "https://www.youtube.com/embed/xxx", title: "标题", desc: "描述", external: true }
```

## 首页展示

首页会自动展示最近更新的 4 篇内容（从 `POSTS` 数组读取），不用额外配置。

## 部署

可以直接丢到 GitHub Pages、Netlify、Vercel 或任意静态托管服务。
