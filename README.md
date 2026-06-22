# trans

维普 Excel → 浙大社 ZD_JATS XML 转换工具。

## 本地运行

```bash
cd code
pip install -r requirements.txt
python converter_gui.py
```

## 命令行

```bash
cd code
python converter.py --xlsx 测试样本/维普数据.xlsx --output-dir ../output --gui
```

## Windows 可执行文件

推送代码到 `main` 分支后，GitHub Actions 会自动在 Windows 环境打包。

1. 打开 [Actions](https://github.com/Essarai/trans/actions)
2. 选择最新的 **Build Windows EXE** 工作流
3. 在 **Artifacts** 中下载 `windows-exe`（`trans-converter.exe`）

也可在 Actions 页手动点击 **Run workflow** 触发构建。
