# Windows 安装包构建说明

目标交付物是一个普通 Windows 安装包：

- 客户下载 `TradeDocGenerator_Setup_*.exe`
- 双击安装
- 桌面图标启动
- 不需要安装 Python

## 资源准备

把业务模板复制到 `resources/templates/`：

- `CPH模板.xlsx`
- `出口危险化学品生产企业符合性声明.docx`
- `产品质量安全责任自负声明.docx`
- `丙烯酸-2-羟基丙酯 厂检单-200kg桶装.docx`
- `丙烯酸-2-羟基丙酯 厂检单吨桶.docx`

如果客户电脑不预装 LibreOffice，二选一：

`resources/LibreOfficePortable/App/libreoffice/program/soffice.exe`

或放官方安装包：

`resources/installers/LibreOffice.msi`

使用 MSI 时，最终安装包会在安装主程序时静默安装 LibreOffice。

MiniMax key 有三种放置方式，优先级从高到低：

- 环境变量 `MINIMAX_API_KEY`
- 用户数据目录 `%LOCALAPPDATA%/TradeDocGenerator/.ai_key`
- 安装资源 `resources/.ai_key`

第一版如果要把自己的 key 打进安装包，可以放 `resources/.ai_key`，但务必先在 MiniMax 后台设置消费上限和余额告警。

## 构建

安装 Inno Setup 6 后，在项目根目录运行：

```powershell
.\build_windows.ps1
```

输出：

- `dist/TradeDocGenerator/`：PyInstaller 应用目录
- `dist-installer/TradeDocGenerator_Setup_1.0.0.exe`：客户安装包

## 数据目录

安装后，上传文件和生成单据写入：

`%LOCALAPPDATA%\TradeDocGenerator`

这样程序即使安装到 `Program Files`，也不会因为目录不可写而失败。
