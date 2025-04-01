# ESP32 烧录工具

一个用 Python 开发的 ESP32 系列芯片多端口并行烧录工具，支持自动检测芯片类型和多固件烧录。

## 功能特点

- 支持多串口并行烧录
- 自动检测芯片型号（ESP32/ESP32-S2/ESP32-S3/ESP32-C3/ESP32-C6/ESP32-P4）
- 支持多固件同时烧录（最多8个）
- 自动保存配置信息
- 支持自动烧录（插入设备自动开始）
- 实时烧录日志显示
- 现代化的图形界面
- 支持高波特率（最高 2000000）
## 支持的芯片
- ESP32
- ESP32-S2
- ESP32-S3
- ESP32-C3
- ESP32-C6
- ESP32-P4
## 使用方法
1. 运行程序：
```
python esp32_flasher.py
 ```

2. 选择串口（支持同时选择多个串口）
3. 选择要烧录的固件文件（.bin）并设置对应的烧录地址
4. 点击"开始烧录"按钮开始烧录过程
## 安装说明

1. 确保已安装 Python 3.x
2. 安装依赖：
```
pip install -r requirements.txt
```
## 开发环境
- Python 3.x
- Windows 10/11
- 依赖库：
  - esptool
  - pyserial
  - tkinter (Python 标准库)
## 许可证
MIT License

## 贡献
欢迎提交 Issue 和 Pull Request！