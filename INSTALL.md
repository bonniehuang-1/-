# 安装指南

## 系统要求

- Python 3.8 或更高版本
- pip（Python包管理器）

## 安装步骤

### 1. 检查Python环境

打开命令行，检查Python是否已安装：

```bash
python --version
```

或

```bash
python3 --version
```

应该显示类似：`Python 3.8.x` 或更高版本

### 2. 检查pip

```bash
python -m pip --version
```

或

```bash
python3 -m pip --version
```

### 3. 安装依赖包

**方法1：使用requirements.txt（推荐）**

```bash
cd production_scheduler
python -m pip install -r requirements.txt
```

**方法2：手动安装**

```bash
python -m pip install pandas>=1.5.0
python -m pip install openpyxl>=3.0.0
python -m pip install numpy>=1.23.0
python -m pip install xlsxwriter>=3.0.0
python -m pip install python-dateutil>=2.8.0
```

### 4. 验证安装

```bash
python -c "import pandas; import openpyxl; import xlsxwriter; print('所有依赖已安装成功！')"
```

## 常见问题

### Q: 提示"pip不是内部或外部命令"

**解决方案**：

1. 使用 `python -m pip` 代替 `pip`
2. 或重新安装Python，确保勾选"Add Python to PATH"

### Q: 安装速度很慢

**解决方案**：使用国内镜像源

```bash
python -m pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple
```

### Q: 权限错误

**解决方案**：

Windows: 以管理员身份运行命令行

Linux/Mac: 使用 `sudo` 或虚拟环境

```bash
python -m pip install --user -r requirements.txt
```

## 使用虚拟环境（推荐）

### Windows

```bash
# 创建虚拟环境
python -m venv venv

# 激活虚拟环境
venv\Scripts\activate

# 安装依赖
pip install -r requirements.txt
```

### Linux/Mac

```bash
# 创建虚拟环境
python3 -m venv venv

# 激活虚拟环境
source venv/bin/activate

# 安装依赖
pip install -r requirements.txt
```

## 快速测试

安装完成后，运行以下命令测试：

```bash
# 生成示例数据
python create_sample_data.py

# 运行系统
python main.py
```

如果一切正常，应该看到系统开始分析并生成报告。

## 离线安装

如果无法联网，可以：

1. 在有网络的机器上下载依赖包：
   ```bash
   pip download -r requirements.txt -d packages
   ```

2. 将packages文件夹复制到目标机器

3. 离线安装：
   ```bash
   pip install --no-index --find-links=packages -r requirements.txt
   ```

## 技术支持

如遇到安装问题，请检查：

1. Python版本是否>=3.8
2. 网络连接是否正常
3. 是否有足够的磁盘空间
4. 防火墙或代理设置

---

安装完成后，请参考 [USER_GUIDE.md](USER_GUIDE.md) 开始使用系统。
