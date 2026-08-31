# Adobe COM MCP Server

用自然语言在 TRAE 里操控 **Photoshop** 和 **Illustrator**。

基于 Windows COM 自动化 + Python MCP，无需 UXP 插件，支持 PS 2020+ 和 AI CS6+。

---

## 功能总览

共 **64 个工具**，覆盖日常设计 + 批量体力活自动化：

### Photoshop — 基础工具（19 个）

| 工具 | 说明 |
|------|------|
| `ps_create_document` | 新建指定像素尺寸的文档 |
| `ps_get_active_info` | 获取当前文档信息（尺寸/图层/模式等） |
| `ps_list_documents` | 列出所有打开的文档 |
| `ps_close_document` | 关闭当前文档（不保存） |
| `ps_add_layer` | 新建空白图层 |
| `ps_duplicate_layer` | 复制当前活动图层 |
| `ps_delete_layer` | 删除当前活动图层 |
| `ps_set_layer_opacity` | 设置图层不透明度（0-100） |
| `ps_add_text_layer` | 添加文字图层（可指定字号/颜色/位置） |
| `ps_fill_layer` | 用指定颜色填充当前图层 |
| `ps_set_foreground_color` | 设置前景色 |
| `ps_apply_gaussian_blur` | 高斯模糊滤镜 |
| `ps_apply_unsharp_mask` | USM 锐化滤镜 |
| `ps_resize_document` | 调整文档像素尺寸（等比或指定） |
| `ps_save_as_png` | 导出 PNG |
| `ps_save_as_jpg` | 导出 JPG |
| `ps_save_as_psd` | 保存 PSD |
| `ps_open_document` | 打开本地图片/PSD 文件 |
| `ps_do_action` | 执行已安装的 Photoshop 动作（Action） |

### Photoshop — 文字图层操作（4 个）

| 工具 | 说明 |
|------|------|
| `ps_list_text_layers` | 列出文档中所有文字图层 |
| `ps_select_all_text_layers` | 一键全选所有文字图层 |
| `ps_replace_text_in_all_layers` | 批量替换所有文字图层中的指定文字 |
| `ps_set_text_color_all` | 批量设置所有文字图层颜色 |
| `ps_set_text_size_all` | 批量设置所有文字图层字号 |

### Photoshop — 图层管理（3 个）

| 工具 | 说明 |
|------|------|
| `ps_list_all_layers` | 列出文档中所有图层（名称/类型/可见性） |
| `ps_select_layer_by_name` | 按名称选中图层 |
| `ps_toggle_layer_visibility` | 切换图层可见性 |
| `ps_batch_rename_layers` | 批量重命名图层（前缀+序号） |

### Photoshop — 批量处理体力活（10 个）★重点

| 工具 | 说明 | 手动耗时 |
|------|------|---------|
| `ps_export_all_layers_to_png` | 一键导出所有图层为单独 PNG | 30层=30轮隐藏/显示/导出 |
| `ps_create_spritesheet` | 图层自动拼成精灵表(Sprite Sheet) | 手动拼图几十分钟 |
| `ps_distribute_layers_evenly` | 多图层均匀分布间距 | 手动计算偏移+逐个移动 |
| `ps_auto_trim` | 自动裁剪透明/白色边缘 | 手动拖裁剪框 |
| `ps_batch_apply_action_to_layers` | 批量给所有图层应用动作 | 逐个选层+执行动作 |
| `ps_smart_object_replace_batch` | 批量替换智能对象内容(mockup) | 逐个打开/替换/保存/关闭 |
| `ps_auto_color_match` | 统一所有图层色调(匹配颜色) | 逐个图层Match Color |
| `ps_auto_center_content` | 自动将内容居中到画布 | 测边界+算偏移+移动 |
| `ps_auto_round_corners` | 批量给图层加圆角(电商图) | 逐个画蒙版/裁切 |
| `ps_auto_layout_strip` | 多图层拼成横向/纵向长条 | 计算尺寸+逐一排列 |
| `ps_export_layer_comps` | 一键导出所有图层组合 | 逐个切换+导出 |
| `ps_batch_resize_folder` | 文件夹批量缩放图片 | 逐张打开/缩放/保存 |
| `ps_batch_watermark` | 文件夹批量加水印 | 逐张加文字+调位置 |
| `ps_social_media_kit` | 一键生成多平台尺寸图 | 每平台单独缩放导出 |
| `ps_contact_sheet` | 生成联系表(缩略图网格) | 手动排版极费时 |
| `ps_auto_gradient_background` | 自动生成渐变背景 | 调渐变编辑器 |
| `ps_smart_replace_text_csv` | CSV批量生成文字版本 | 每条数据手动改字 |
| `ps_auto_layout_cards` | 自动网格排版多图层卡片 | 手动排版定位 |
| `ps_create_gif_from_layers` | 图层→GIF动画 | 手动逐帧设置 |
| `ps_extract_color_palette` | 提取图片主色调色盘 | 手动吸取+记录 |

### Illustrator（11 + 4 个）

| 工具 | 说明 |
|------|------|
| `ai_create_document` | 新建指定尺寸的文档（点） |
| `ai_get_active_info` | 获取当前文档信息 |
| `ai_list_documents` | 列出所有打开的文档 |
| `ai_close_document` | 关闭当前文档 |
| `ai_add_rectangle` | 画矩形（可指定填充色） |
| `ai_add_ellipse` | 画椭圆（可指定填充色） |
| `ai_add_polygon` | 画正多边形（可指定填充色） |
| `ai_add_text` | 添加文字（可指定字号/颜色/位置） |
| `ai_save_as_ai` | 保存为 .ai |
| `ai_export_svg` | 导出 SVG 矢量图 |
| `ai_export_png` | 导出 PNG（透明背景） |
| `ai_batch_replace_text` | 批量查找替换所有文字框内容 |
| `ai_export_all_artboards` | 一键导出所有画板为单独文件 |
| `ai_auto_layout_grid` | 自动网格排列所有对象 |
| `ai_align_objects` | 多对象对齐（左/右/顶/底/居中） |

---

## 安装步骤

### 1. 环境要求

| 组件 | 支持版本 | 已验证版本 |
|------|---------|-----------|
| 操作系统 | Windows（需支持 COM 自动化） | Windows 11 |
| Photoshop | **2020 及以上**（CS6+ 理论上可用） | PS 2020 (v21.0.1) |
| Illustrator | **CS6 及以上** | AI 2023 (v27.0.0) |
| Python | **3.10 及以上** | — |
| Python 包 | `pywin32` + `mcp`（`mcp<2`） | — |

> ⚠️ **版本说明**：Photoshop 通过 `DoJavaScript` 执行 ExtendScript 实现自动化，这是跨版本最稳定的方式，PS 2020+ 全部适用；Illustrator 通过 COM 对象直接调用 API，AI CS6+ 均可。若你的版本较新（如 PS 2024/2025、AI 2024），接口向下兼容，正常可用。

### 2. 安装依赖

```bash
pip install pywin32 "mcp<2"
```

### 3. 在 TRAE 中添加 MCP 服务器

1. 打开 TRAE
2. 左下角点击 **设置** → **MCP**
3. 点击 **添加 MCP 服务器**
4. 选择 **命令行（stdio）** 类型
5. 填写配置：

   - **名称**：`adobe-com-mcp`
   - **命令**：`python`
   - **参数**：`C:\Users\你的用户名\Documents\adobe-com-mcp\server.py`
   - **工作目录**：`C:\Users\你的用户名\Documents\adobe-com-mcp`

   （把路径替换成你实际的 server.py 位置）

6. 保存并启用

### 4. 验证

打开 Photoshop，然后在 TRAE 聊天框说：

> 新建一个 800x800 的 Photoshop 文档

如果 PS 自动创建了文档 → 配置成功！

---

## 使用示例

### Photoshop 示例

**新建画布并加文字：**
> "帮我建一个 1920x1080 的 PS 文档，加个标题文字叫'你好世界'，字号 72，放在左上角"

**做一张海报底图：**
> "新建 1080x1920 的竖版画布，背景色用 #FF6B35，加一个文字图层写'SUMMER SALE'，白色 96 号字"

**修图操作：**
> "给当前图层加 3 像素高斯模糊，然后把不透明度调到 70%"

**导出成品：**
> "把这张图导出成 PNG，存到桌面"

### Illustrator 示例

**画个 Logo 雏形：**
> "在 AI 里新建 800x600 的画布，画一个蓝色正六边形，中间加文字'LOGO'"

**导出矢量图：**
> "把当前 AI 文档导出成 SVG"

### 批量体力活示例 ★重点

**导出所有图层为单独 PNG（30层只需1句话）：**
> "把当前文档所有图层导出成单独的 PNG，存到桌面 layers_export 文件夹"

**生成精灵表（游戏开发）：**
> "把所有图层拼成一张精灵表，自动排列网格"

**批量替换智能对象（Mockup 批量生成）：**
> "把 D:\mockup_images 文件夹里的图片依次替换到文档中的智能对象图层"

**统一色调：**
> "以当前活动图层为参考，统一所有图层的色调"

**批量加圆角（电商商品图）：**
> "给当前文档所有可见图层加 20px 圆角"

**AI 批量查找替换：**
> "把 AI 文档里所有文字框中的'2024'替换成'2025'"

**AI 一键导出所有画板：**
> "把 AI 文档所有画板导出成 PNG，存到桌面"

**图层均匀分布：**
> "把我选中的这些图层水平均匀分布间距"

**拼长图：**
> "把当前文档图层排成纵向长条，间距 10px"

---

## 工作原理

1. MCP 服务器（Python 进程）通过 COM 接口连接 Photoshop / Illustrator
2. Photoshop 操作通过 `DoJavaScript` 执行 ExtendScript（最稳定的 PS 自动化方式）
3. Illustrator 操作通过 COM 对象直接调用 API
4. 遇到 "应用程序忙" 时自动重试（最多 40 次，间隔 250ms）
5. 每个线程独立持有 COM 对象，避免跨线程 marshalling 问题

---

## 常见问题

**Q: PS 一直报"应用程序正在使用中"怎么办？**
A: 检查 PS 是否有模态对话框（欢迎屏幕、字体缺失提示等），关掉即可。也可以重启 PS 后再试。

**Q: 可以同时操控 PS 和 AI 吗？**
A: 可以，一个 MCP 服务器同时支持两个软件，工具名分别以 `ps_` 和 `ai_` 开头。

**Q: 支持哪些 PS 版本？**
A: 测试通过 PS 2020 (v21.0.1)。理论上 CS6 及以上都可用。

**Q: 支持哪些 AI 版本？**
A: 测试通过 AI 2023 (v27.0.0)。理论上 CS6 及以上都可用。

**Q: 需要打开 PS/AI 才能用吗？**
A: 如果软件没开，MCP 会自动启动它。但建议提前打开，响应更快。
