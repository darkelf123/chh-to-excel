# CHH摸鱼助手 (excel)

> 把 [Chiphell](https://www.chiphell.com/) 论坛伪装成 Excel 表格。上班摸鱼，老板路过看不出来。

![版本](https://img.shields.io/badge/version-1.0-green)
![License](https://img.shields.io/badge/license-MIT-blue)
![Tampermonkey](https://img.shields.io/badge/Tampermonkey-required-orange)

---

## 📖 简介

在办公室里想刷 Chiphell 又怕被发现？装这个脚本，按 `Ctrl+Shift+X` 一键把整个论坛变成 Excel：

- 顶部绿色 Excel 标题栏 (`Hardware_Discussion_2026.xlsx`)
- 帖子正文变成单元格表格（白底、细边框、22px 行高）
- 公式栏显示可点击的面包屑导航
- 底部 Sheet 标签对应分页器
- **附件图片默认全部隐藏**（老板最安全的摸鱼模式），点逐图按钮或全局按钮一键显示

老板扫一眼：哟，又在做表。完美。

---

## 🖼️ 截图

> _装好脚本后，截两张图替换这里的占位符：_

| 列表页（摸鱼模式） | 帖子页（图片默认隐藏） |
|---|---|
| ![列表页](assets/screenshot-list.png) | ![帖子页](assets/screenshot-thread.png) |

---

## ✨ 核心特性

- 🎨 **真·Excel 视觉** — Office Green 配色（`#217346`）、DengXian 字体、单元格细边框、`fx` 公式栏、`Sheet` 翻页
- 🧭 **面包屑公式栏** — 鼠标移入单元格显示文本，移出恢复面包屑链接
- 📑 **Sheet 翻页** — Discuz! 翻页器变成 Sheet 标签，当前页加粗绿色
- 🖼️ **图片默认隐藏**（移植自 NGA）— 进帖子立刻看到的是白底表格，按需逐图/全部显示
- ⌨️ **快捷键** — `Ctrl+Shift+X` 一键切换摸鱼/正常模式
- 💾 **状态记忆** — `localStorage` 记住你的摸鱼偏好（`chh_excel_mode`）

---

## 🙏 灵感来源 & 致谢

本脚本**不是从零原创**，核心机制借鉴了两个优秀的摸鱼脚本。在此明确声明：

### 1. [S1摸鱼助手(excel)](https://greasyfork.org/zh-CN/scripts/557042-s1%E6%91%B9%E9%B1%BC%E5%8A%A9%E6%89%8B-excel) — **by Gemini**

**借鉴深度：95% 移植（核心架构 + 全部 CSS + 全部 JS 逻辑）**

这是本脚本的 **source of truth**。原脚本把 [Stage1st 论坛 (bbs.saraba1st.com)](https://bbs.saraba1st.com/2b/) 伪装成 Excel。本脚本对其做了以下改动：

| 项 | 原 (S1) | 本脚本 (CHH) |
|---|---|---|
| `@match` 域名 | `bbs.saraba1st.com/2b/*` | `www.chiphell.com/{forum-*,thread-*,forum.php}` |
| 伪装的文件名 | `Business_Report_2025.xlsx` | `Hardware_Discussion_2026.xlsx` |
| 公式栏首页文字 | `Stage1st` | `Chiphell` |
| `localStorage` key | `s1_excel_mode` | `chh_excel_mode` |
| 图片隐藏 | ❌ 无 | ✅ 移植自 NGA，见下 |

**未改动**：HTML 注入模板、CSS 隐藏规则（`#toptb #hd #nv_ph` 等 14 个 Discuz! 标准选择器）、JS 主体逻辑（`enableExcel / disableExcel / toggleMode / initBreadcrumb / initSheetPagination / updateFormulaBar`）。

**为什么几乎原样可用**：Chiphell 和 Stage1st 都使用 **Discuz! X3.4 标准模板**，所有 ID/类选择器（`#pt #threadlisttableid #postlist .plhin .pls .plc .t_f .pg` 等）在两个论坛都 1:1 命中，已实测验证。

### 2. [NGA优化-摸鱼体验](https://greasyfork.org/zh-CN/scripts/393991-nga%E4%BC%98%E5%8C%96%E6%91%B8%E9%B1%BC%E4%BD%93%E9%AA%8C) — **by kisshang1993**

**借鉴深度：核心机制移植（约 60 行 JS 重写为 vanilla JS）**

本脚本的 **图片默认隐藏 + 逐图按钮 + 全局开关** 功能直接移植自这个脚本。NGA 原版的实现是 jQuery 风格的，本脚本用原生 JS 重写并适配 chiphell 的 `.t_f img.zoom`（Discuz! 附件图标准类）。

原 NGA 核心逻辑（节选自 Greasy Fork 源码）：

```js
// 原 NGA 实现 (jQuery)
$('img').each(function () {
    if (!class.includes('smile')) {  // 跳过表情
        $(this).addClass('hld__img-postimg')  // 标记可隐藏
        let $btn = $('<button class="switch-img" style="display:none">图</button>')
        $btn.on('click', function () {
            $(this).prev('img').toggle()
            $(this).text($(this).prev('img').is(':hidden') ? '图' : '隐藏')
        })
        $(this).after($btn)
    }
})
```

本脚本移植版（原生 JS + 默认隐藏 + chiphell 适配）：

```js
// 本脚本 (vanilla JS)
const images = document.querySelectorAll('.t_f img.zoom:not([data-hld-injected])');
images.forEach(img => {
    img.dataset.hldInjected = '1';
    img.classList.add('hld__img-hidden');  // 默认隐藏
    const btn = document.createElement('button');
    btn.className = 'excel-img-toggle';
    btn.textContent = '📷 显示图片';
    btn.addEventListener('click', function() {
        img.classList.toggle('hld__img-hidden');
        btn.textContent = img.classList.contains('hld__img-hidden') ? '📷 显示图片' : '🚫 隐藏图片';
    });
    img.parentNode.insertBefore(btn, img.nextSibling);
});
```

**与 NGA 的差异**：
- 默认隐藏（NGA 默认 `hideImage: false`，用户主动开启；本脚本反之）
- 增加全局按钮 + 状态联动逻辑
- 跳过规则改用 `.t_f img.zoom` 选择器，自动避开头像/UI 图标

### 本脚本完全原创的部分

- ✅ Chiphell 适配（match 规则 + 文字本地化 + 主题色适配）
- ✅ **默认隐藏策略**（NGA 是用户主动开启，本脚本相反）
- ✅ **双重按钮 UX**（逐图按钮 + 右上角全局按钮 + 状态联动）
- ✅ **`data-hld-injected` 重复注入防护**（Ctrl+Shift+X 切换时不重复添加按钮）
- ✅ **CSS 配色抽出为常量**（`COLOR_PRIMARY / COLOR_PRIMARY_DARK / COLOR_HOVER_BG`），便于以后改主题

---

## ⚙️ 工作原理（开发者向）

### 整体架构

```
┌─────────────────────────────────────────┐
│  CSS 重写（颜色 + 字号 + 边框 + 隐藏）   │  ← 90 行
├─────────────────────────────────────────┤
│  注入 Excel UI 头部 + Sheet 栏 + 状态栏  │  ← HTML 模板字符串
├─────────────────────────────────────────┤
│  JS 逻辑                                │
│  ├─ toggleMode()    切换摸鱼状态         │
│  ├─ enableExcel()   注入 + 隐藏原界面     │
│  ├─ disableExcel()  恢复原界面           │
│  ├─ initBreadcrumb() 克隆面包屑到公式栏  │
│  ├─ initSheetPagination() 翻页→Sheet    │
│  ├─ updateFormulaBar() 鼠标交互切换       │
│  └─ initImageHider() 隐藏图片 + 注入按钮  │  ← 移植自 NGA
└─────────────────────────────────────────┘
```

### Discuz! 兼容性

依赖的 Discuz! 标准 ID/类（chiphell 全部 1:1 命中，已实测）：

| 选择器 | 用途 |
|---|---|
| `#toptb #hd #nv_ph #scbar #ft #scrolltop #pt` | 顶部/底部/导航/面包屑（hide） |
| `#threadlisttableid` | 列表表格 |
| `#postlist .plhin .pls .plc .t_f` | 帖子正文结构 |
| `.pg` | 翻页器 |
| `.t_f img.zoom` | 附件图（hide 目标） |

### localStorage

| Key | Value | 说明 |
|---|---|---|
| `chh_excel_mode` | `'true' / 'false'` | 摸鱼状态记忆 |

---

## 📥 安装使用

### 前置要求

- Chrome / Edge / Firefox 浏览器
- [Tampermonkey](https://www.tampermonkey.net/) 扩展

### 安装步骤

1. **装 Tampermonkey**：[Chrome Web Store](https://chrome.google.com/webstore/detail/tampermonkey/dhdgffkkebhmkfjojejmpbldmpobfkfo) 或 [Edge Add-ons](https://microsoftedge.microsoft.com/addons/detail/tampermonkey/iikmkjmpaadaobahmlepeloendndfphd)
2. **新建脚本**：Tampermonkey 面板 → "添加新脚本"
3. **粘贴代码**：把 [`chh_excel.user.js`](chh_excel.user.js) 文件内容粘贴进去保存
4. **访问论坛**：打开 [chiphell.com](https://www.chiphell.com/)，右上角出现 "摸鱼模式" 按钮即安装成功

### 使用方法

| 操作 | 效果 |
|---|---|
| 点 "摸鱼模式" 按钮 | 进入摸鱼模式 |
| `Ctrl+Shift+X` | 切换摸鱼/正常模式 |
| 点 "🖼️ 全部显示" 按钮 | 显示当前帖子所有图片 |
| 点某个图后面的 "📷 显示图片" | 单独显示那张图 |
| 点 "🚫 隐藏图片" | 单独隐藏那张图 |
| 关闭摸鱼模式 | 全部恢复正常 |

---

## 🔧 自定义配置

脚本顶部有几个常量可以自定义（修改后保存即可生效）：

```js
const COLOR_PRIMARY      = '#217346';   // Excel 经典绿
const COLOR_PRIMARY_DARK = '#104e2b';   // 深绿 (hover)
const COLOR_HOVER_BG     = '#e6f2ea';   // 浅绿 hover 背景
const COLOR_TEXT         = '#000';      // 文字色
const COLOR_BORDER       = '#e1e1e1';   // 边框色
```

想换主题色？比如改成经典 Office 2003 蓝：
```js
const COLOR_PRIMARY      = '#1F497D';
const COLOR_PRIMARY_DARK = '#0F2A4D';
const COLOR_HOVER_BG     = '#D8E4F0';
```

---

## 🗺️ 开发路线图

- [ ] 支持 `Ctrl+Shift+H` 一键隐藏/显示所有图片（快捷键）
- [ ] 支持更多 Discuz! 论坛（理论上零改动可用，需要测试）
- [ ] 列表页支持"只看楼主"快捷开关
- [ ] 添加暗色模式（Excel 2016+ 黑色主题）
- [ ] 添加图片懒加载（按需加载隐藏图片）
- [ ] 支持用户自定义文件名（不写死在脚本里）

---

## 📜 License

MIT License — 与原 S1 摸鱼助手一致。

如果你修改或转发，请保留对原作者 Gemini (S1 脚本) 和 kisshang1993 (NGA 脚本) 的致谢。

---

## 🐛 反馈

- **Issues**: [GitHub Issues](../../issues)
- **原 S1 脚本反馈**: [Greasy Fork](https://greasyfork.org/zh-CN/scripts/557042-s1%E6%91%B9%E9%B1%BC%E5%8A%A9%E6%89%8B-excel)
- **原 NGA 脚本反馈**: [Greasy Fork](https://greasyfork.org/zh-CN/scripts/393991-nga%E4%BC%98%E5%8C%96%E6%91%B8%E9%B1%BC%E4%BD%93%E9%AA%8C)

---

<p align="center">
  <sub>摸鱼愉快 🚀 — 记得别被发现</sub>
</p>