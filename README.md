# Amazon Invoice 合并工具

将多份 Amazon Invoice Excel 文件合并为一张总表，自动按 Tracking ID 透视汇总各项费用，生成含税/不含税合计列。

## 功能说明

导入一个或多个 Amazon Invoice (.xlsx / .xls) 文件后，程序自动识别格式，按 Tracking ID 将各项 Charge Code（基础运费、体积附加费、偏远地区费等）展开为独立列，O 列为不含税合计（SUM H:N），P 列为含税合计（O × 1.20，税率 20%）。

## 获取 Windows 可执行文件（.exe）

### 方式一：GitHub Actions 云端自动打包（推荐）

将代码推送到 GitHub 后，云端 Windows 机器会自动打包，约 5 分钟完成。

```bash
gh auth login          # 登录 GitHub，跟着提示走
cd excel_merger
gh repo create amazon-invoice-merger --public --push --source .
```

推送完成后，进入 GitHub 仓库页面，点击顶部 Actions 标签，选择最新一条构建记录，向下滚动找到 Artifacts，下载 Amazon合并工具-Windows.zip，解压即得 Amazon合并工具.exe。

以后每次修改代码后推送一下，自动重新打包，无需任何手动操作。

### 方式二：在 Windows 上本地打包

把整个 `excel_merger` 文件夹拷贝到 Windows 电脑，安装好 Python（从 python.org 下载，安装时勾选 "Add to PATH"），然后双击 `build.bat`，约 3 分钟后 `dist\Amazon合并工具.exe` 即生成完毕。之后只需将这一个 .exe 文件复制到任意 Windows 电脑，双击即可直接运行，无需安装 Python 或任何依赖。

## 输出列说明

| 列 | 字段 | 说明 |
|---|---|---|
| A | Tracking ID | 运单号 |
| B | To Postcode | 目的地邮编 |
| C | Reference | 参考单号 |
| D | Billable weight (kg) | 计费重量 |
| E~G | Length / Width / Height | 尺寸 (cm) |
| H | Base charge | 基础运费 |
| I | High Cube Surcharge | 体积附加费 |
| J | Delivery Area Surcharge | 偏远地区费 |
| K | Misdeclaration Handling Charge | 错误申报费 |
| L | Base Rate Adjustment | 费率调整 |
| M | Non-Conveyable Surcharge | 不可传送附加费 |
| N | Additional Handling Fees: Girth | 附加处理费（围长） |
| O | 合计（不含税） | SUM(H:N) |
| P | 发票金额（含税） | O × 1.20 |
| Q | 备注 | 手动填写 |
