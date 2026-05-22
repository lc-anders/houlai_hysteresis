# houlai_hysteresis 滞回曲线处理与抗震评价程序

## 程序定位

`houlai_hysteresis` 是一个基于 Python + Streamlit 的滞回曲线处理与抗震评价程序，主要用于低周反复加载试验中的位移—荷载/反力数据处理。

本版本已统一去除早期临时命名，避免与已有商业软件或第三方工具名称产生混淆。

## 主要功能

- 上传 Excel / CSV / TXT 数据文件；
- 选择位移列与荷载/反力列；
- 停顿点修正、隔行取数、平滑处理；
- 加载级位移峰值法分圈；
- 自动计算滞回环面积、累计耗能、割线刚度、等效黏滞阻尼系数、残余变形；
- 输出位移峰值骨架、荷载峰值骨架和包络骨架；
- 延性系数计算；
- 支持中文/英文图表输出；
- 支持 Excel、PNG/SVG/PDF、GIF 和 ZIP 批量导出。

## 本地运行

### 方式一：双击启动

Windows 用户可直接双击：

```text
本地一键启动_houlai_hysteresis.bat
```

如果 Python 路径不是 `D:\Software\Python\Python310\python.exe`，请使用：

```text
本地一键启动_通用Python.bat
```

### 方式二：命令行运行

```powershell
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

或使用你的固定 Python 路径：

```powershell
D:\Software\Python\Python310\python.exe -m pip install -r requirements.txt
D:\Software\Python\Python310\python.exe -m streamlit run app.py
```

## Streamlit Community Cloud 在线部署

GitHub 仓库根目录至少需要包含：

```text
app.py
requirements.txt
.streamlit/config.toml
```

部署页面填写：

```text
Repository: 你的用户名/你的仓库名
Branch: main
Main file path: app.py
```

建议 App URL 使用：

```text
houlai-hysteresis
```

或：

```text
houlai-hysteresis-lc
```

## 建议 GitHub 仓库名

推荐仓库名：

```text
houlai_hysteresis
```

或：

```text
houlai-hysteresis
```

## 注意事项

1. 程序仅用于试验数据处理与科研分析。
2. 在线部署后，上传文件会在平台临时环境中处理，正式试验原始数据建议谨慎上传第三方平台。
3. 中文图片如乱码，程序中字体优先选择 `Microsoft YaHei` 或 `SimHei`。
4. 对斗栱、榫卯、木结构节点等存在滑移、捏拢、摩擦特征的滞回数据，自动分圈结果仍建议结合试验记录人工复核。
