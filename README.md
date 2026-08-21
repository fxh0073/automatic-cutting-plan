# Automatic Cutting Plan

面向型材下料工艺卡的开源下料方案生成器。

## 当前状态

公开依赖版已经可以从示例工艺卡生成结果。仓库不包含真实工艺卡、客户资料、输出文件、运行时和构建产物；公开发布前仍需完成公司授权和代码审查。

## 功能目标

- 读取包含“锯床”或“下料”工作表的 `.xlsx` 工艺卡；
- 解析部件、材质、规格尺寸和数量；
- 按原料定尺、锯缝和取整步长生成型材切割组合；
- 输出下料方案、切割明细和数量校核工作表；
- 支持 Windows 本地运行；独立版发布包作为 GitHub Release 附件单独发布。

本项目使用公开的 [ExcelJS](https://github.com/exceljs/exceljs) 读写 XLSX；不依赖私有的工作区工具包。ExcelJS 及其依赖按项目许可证使用。

## 两种使用方式

### 独立版（现场使用）

现场电脑使用 GitHub Releases 中的“自动下料程序_独立版”压缩包。独立版自带 Node.js 运行时和程序所需文件，目标电脑不需要安装 Excel、Node.js 或 Codex。请先完整解压文件夹，再双击 `自动下料程序.exe`，不要单独复制 exe 文件。

完整操作说明见：[独立版使用说明](docs/独立版使用说明.md)。

当前源码仓库不提交 exe、运行时、ZIP 包或真实工艺卡；这些内容应在发布 Release 时单独整理并再次检查。

独立版可以在有 Node.js 和 .NET Framework 开发环境的 Windows 电脑上通过以下脚本重新打包：

```powershell
.\scripts\build_standalone.ps1 -OutputDirectory .\dist\自动下料程序_独立版
```

脚本会把公开依赖、Node.js 运行时、ExcelJS 引擎和 Windows 启动器放入同一个目录，并检查发布包中没有私有工作区依赖。目标电脑不需要安装 Node.js。

### 源码开发运行

- Node.js 18 或更高版本；
- `npm install` 安装公开依赖；
- 输入文件必须使用脱敏后的示例，不要提交真实项目工艺卡。

## 运行

```powershell
npm install
node src/build_cutting_plan.mjs --input .\examples\下料工艺卡_示例.xlsx --output .\outputs\示例下料方案.xlsx
npm test
```

## 输入格式

默认识别 `锯床` 工作表，也支持名称中包含“锯床”或“下料”的工作表。表头应包含：序号、部件名称、材质、规格尺寸、数量；图号和备注为可选字段。

## 许可

本仓库包含 MIT License。

## 范围说明

当前项目生成的是下料方案和校核表，不直接向锯床下发指令。实际联机控制需要结合具体设备型号、通信协议和安全联锁另行开发。
