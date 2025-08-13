# 数据校验加载项 v1.0.3

一个基于Python的Excel数据校验工具，提供图形化界面进行数据验证和质量检查。

## 功能特性

- **Excel文件处理**: 支持.xlsx和.xls格式文件的读取和验证
- **数据校验**: 提供多种数据验证规则，包括数据类型、格式、范围检查
- **批量处理**: 支持批量文件处理和结果汇总
- **错误报告**: 生成详细的错误日志和校验报告
- **用户友好界面**: 直观的图形化操作界面
- **结果导出**: 支持将校验结果导出为Excel或CSV格式

## 技术栈

- **Python 3.10+**
- **PyQt5**: 图形用户界面
- **pandas**: 数据处理和分析
- **openpyxl**: Excel文件操作
- **PyInstaller**: 应用程序打包

## 安装和运行

### 环境要求

- Python 3.10 或更高版本
- Windows 10/11 操作系统

### 安装依赖

```bash
pip install -r requirements.txt
```

### 运行应用

```bash
python main.py
```

### 打包应用

```bash
pyinstaller main.spec
```

## 使用方法

1. **启动应用**: 双击main.exe或运行python main.py
2. **选择文件**: 点击"选择文件"按钮选择要校验的Excel文件
3. **设置规则**: 配置数据校验规则
4. **执行校验**: 点击"开始校验"按钮
5. **查看结果**: 查看校验结果和错误报告
6. **导出结果**: 将结果导出为Excel或CSV文件

## 项目结构

```
project3.0 - 20250810v2T - v1.0.3/
├── main.py              # 主程序入口
├── ui.py               # 用户界面模块
├── checker.py          # 数据校验核心模块
├── requirements.txt    # 项目依赖
├── main.spec          # PyInstaller打包配置
├── app_icon.ico       # 应用图标
└── README.md          # 项目说明文档
```

## 开发说明

### 添加新校验规则

在`checker.py`中添加新的校验函数，然后在UI模块中集成。

### 修改界面

使用`ui.py`中的PyQt5代码进行界面调整。

## 版本历史

- **v1.0.3**: 初始版本，包含基本的数据校验功能

## 许可证

MIT License

## 贡献

欢迎提交Issue和Pull Request来改进这个项目。