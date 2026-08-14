# 报销助手（Tauri）

轻量桌面应用：**托盘驻留 + 全局快捷键唤起（⌘/Ctrl+Shift+R）+ 关窗不退出**，
跨平台（macOS / Windows / Linux）。全程本地处理，无大模型、无联网上传。

## 功能

- **发票签字**：批量给电子发票 PDF 加电子签、写归档标识、规范重命名
  （`公司-金额-dzfp票号-已签字.pdf`）。
- **发票明细 Excel**：按报销官方订单表模板结构输出
  （项目名称/财务项目码/发票号/物品名称/物品规格/供应商编码/供应商名称/单位/单价/数量），
  明细从发票 PDF 文本层自动提取，空值填"无"，发票号列为文本格式。
- **实物证据 Word**：发票号 / 公司 / 税号 / 采购日期 + 实物照片，
  宋体 12pt、1.5 倍行距（`公司-金额-dzfp票号-实物证据.docx`）。
- **一键处理**：签字 + 明细 Excel + 实物证据 Word + AI 校验一次完成；
  补拖照片自动更新 Word/Excel 并重新校验。
- **四槽位校验**：PDF / Excel / Word / 照片逐项核对，通过显示 ✓。
- **项目台账**：项目名称 + 经费代码记忆在下拉框里，多项目切换。

## 架构

```
tauri-app/
├── ui/                       # 前端（纯静态 HTML/CSS/JS，无构建步骤）
│   ├── index.html / style.css
│   ├── app.js                # 界面逻辑
│   └── tauri-bridge.js       # window.invoiceApp 适配层 → Tauri IPC
├── .github/workflows/build.yml  # GitHub Actions 四平台在线打包
└── src-tauri/
    ├── src/lib.rs            # Rust 后端：托盘/快捷键/对话框/任务编排/项目台账
    ├── helpers/              # Python 处理引擎（源码）
    │   ├── add_invoice_signature.py   # 发票签字 + 文本层字段/明细提取
    │   ├── invoice_summary.py         # 发票明细 Excel（官方模板）
    │   ├── invoice_evidence.py        # 实物证据 Word
    │   ├── verify_package.py          # 四槽位校验
    │   └── reagent_report.py          # 试剂耗材汇总（CLI 保留）
    ├── tauri.conf.json
    └── capabilities/
```

引擎解析顺序：`helpers/reimburse-helper` 打包二进制 → 系统 Python（需
PyMuPDF / Pillow / openpyxl / python-docx）。有打包二进制时无需装 Python。

## 在线打包（GitHub Actions）

推送代码或打 `v*` tag 自动构建四个平台安装包：

- macOS arm64（Apple 芯片）+ macOS x64（Intel）：`.dmg`
- Windows x64：`.msi` / `.exe`（NSIS）
- Linux x64：`.deb` / `.AppImage`

产物在 Actions 的 Artifacts 下载；打 tag（如 `v0.2.0`）会自动发布为
GitHub Release。

## 本地开发运行

前置：Rust 工具链（rustup）、Node.js。

```bash
npm install --registry=https://registry.npmmirror.com
npm run dev
```

## 本地打包

```bash
# 先打包 Python 引擎（见下），再构建应用
npm run build
```

产物：`src-tauri/target/release/bundle/`（dmg / msi / deb 等按平台）。

## 重新打包 Python 引擎

修改了 `add_invoice_signature.py` / `invoice_summary.py` 等脚本后：

```bash
cd src-tauri/helpers
python3 -m venv --system-site-packages /tmp/pyinstaller-venv
/tmp/pyinstaller-venv/bin/pip install pyinstaller
/tmp/pyinstaller-venv/bin/pyinstaller --onedir --name reimburse-helper \
  --exclude-module pandas --exclude-module scipy --exclude-module matplotlib \
  --exclude-module contourpy --exclude-module kiwisolver \
  --exclude-module fontTools --exclude-module dateutil --exclude-module pytz \
  --add-data "add_invoice_signature.py:." \
  --add-data "reagent_report.py:." \
  --add-data "invoice_summary.py:." \
  --add-data "invoice_evidence.py:." \
  --add-data "verify_package.py:." \
  helper_cli.py
cp -r dist/reimburse-helper .
rm -rf dist build
```

Windows 上 `--add-data` 分隔符用 `;`。
引擎不内置任何电子签：使用者首次启动时在界面里选择自己的签名图片。

## macOS 安装与 Gatekeeper 说明

应用为 adhoc 签名（未购买 Apple Developer ID，无法公证），macOS 首次打开会拦截，
提示"无法验证开发者"或"已损坏"。**安装后执行一次**：

```bash
cp -R "src-tauri/target/release/bundle/macos/报销助手.app" /Applications/
xattr -cr "/Applications/报销助手.app"   # 清除隔离属性，之后可正常双击
open "/Applications/报销助手.app"
```

从 dmg 重新安装后若再次被拦截，只需重新执行 `xattr -cr "/Applications/报销助手.app"`。
正式对外分发时再考虑 Developer ID 签名 + 公证。

## 隐私

- 应用包不内置任何个人签名或数据；发票、照片、签名全程本机处理，不上传。
- 项目台账存于系统应用配置目录（`projects.json`），仅在安装者本机。

## 全局快捷键

`Command/Ctrl + Shift + R` 唤起主窗口；关闭窗口时隐藏到菜单栏托盘，不退出。
