# TC Shell Listener

**TC Shell Listener** 是一个辅助工具，用于监控 Windows 文件资源管理器（Explorer）的文件打开动作，自动用 Total Commander（TC）接管文件/文件夹的打开，并自动将原有 Explorer 窗口移出屏幕或关闭，实现 Explorer->Total Commander 的无缝跳转。

## 功能特性

- 自动检测当前选择的文件/文件夹
- 自动调用 Total Commander 打开目标
- 一键隐藏并关闭 Explorer 窗口，视觉无打扰
- 首次运行支持自动查找已运行的 TC，也支持手动选择 TC 路径，后续自动记录配置
- 配置文件记录 TC 路径，无需每次选择

## 使用说明

1. **准备工作**
   - 请确保 [Total Commander](https://www.ghisler.com/) 已正确安装并可正常启动。
   - 本程序需 .NET Framework 4.5 或以上（或 .NET Core 依赖 WinForms 支持）。
2. **启动程序**
   - 双击启动 `TC Shell Listener.exe` 或运行编译后的应用。
   - 首次运行如果未检测到已开启的 Total Commander，会弹出对话框要求手动选择 TC 可执行文件（例如 `TOTALCMD64.EXE`）。
   - 选择路径后，系统会自动记录到根目录下的 `tc_config.txt`，后续无需再次设置。
3. **自动管理资源管理器**
   - 此后每次尝试用 Windows 文件资源管理器打开文件或目录时，程序会自动捕获并转交给 Total Commander 打开，并把 Explorer 窗口移出屏幕或自动关闭。
4. **退出**
   - 直接关闭程序即可。

## 配置文件说明

- 程序根目录下 `tc_config.txt`，内容仅为 Total Commander 的完整路径。
- 若 TC 路径改变，删除该文件即可触发重新选择。

## 常见问题

- **Q:** 资源管理器窗口被移走无法正常操作怎么办？  
  **A:** 请通过 Total Commander 操作文件，或重启本程序。
- **Q:** 新换了 Total Commander 路径如何重选？  
  **A:** 删除目录下的 `tc_config.txt` 后重启程序即可重新选择路径。
- **Q:** 高权限运行时需管理员身份吗？  
  **A:** 通常不需要。但如遇报错可尝试“以管理员身份运行”本工具。

## 参与开发

如有建议或问题欢迎提交 [Issues](https://github.com/your-repo/issues)。

---

**免责声明**：  
本软件为个人辅助编写，使用存在一定系统操作风险，请自行备份重要数据和配置，出现问题请自行承担风险。
