# 参与贡献

感谢你愿意改进 DocTabs。提交代码前请先看下面几条。

## 提交 Issue

请尽量附上：

- Windows 版本（如 Windows 11 23H2）
- Office 版本（Word / Excel → 文件 → 账户 → 关于）
- DocTabs 版本（托盘菜单或 `DocTabs.exe` 属性 → 详细信息）
- 复现步骤与期望结果
- 相关日志：`%LOCALAPPDATA%\DocTabs\DocTabs.log`

## 提交 Pull Request

1. 从 `main` 开分支，例如 `fix/tab-color`。
2. 保持单文件结构：除非确有必要，不要把 `DocTabs.cs` 拆成多个文件。
3. 不引入第三方依赖；目标运行时是 .NET Framework 4.7.2，语言级别按 C# 5 编写。
4. 运行 `build.bat` 确认能编译，并在本机 Word / Excel 上验证改动。
5. PR 描述写清楚：问题现象、根因、修改点、验证方式。

## 代码约定

- 4 空格缩进，保持现有命名与注释风格。
- Win32 调用统一放在 `Native` 类。
- 新增日志使用现有 `Log` 类，避免高频刷日志。
- 不写死本机路径、账号或安装目录。

## 版本与发布

- 版本号在 `DocTabs.cs` 的 `AssemblyVersion` / `AssemblyFileVersion` 中维护。
- 每次发布同步更新 `CHANGELOG.md`。
- 发布时在 GitHub Releases 上传 `stage\DocTabs.exe`，标签用 `v1.5.x`。
