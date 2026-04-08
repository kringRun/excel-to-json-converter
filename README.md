# Excel转JSON转换器

## 功能说明
将GVL输入输出表格（Excel格式）转换为JSON格式，自动处理：
- 去掉变量名中的 DI_/DO_ 前缀
- 去掉地址中的 % 符号
- 保留中文名称和地址

## 使用方法

### 方法一：图形界面（推荐）
1. 双击运行 `Excel转JSON工具.exe`
2. 点击"浏览..."选择Excel文件
3. 自动设置JSON输出路径（可修改）
4. 点击"开始转换"按钮
5. 查看转换结果

### 方法二：命令行
```
Excel转JSON工具.exe <excel文件路径> [json输出路径]
```

示例：
```
Excel转JSON工具.exe "D:\data\GVL_输入输出.xlsx" "D:\output\result.json"
```

## 支持的Excel格式
- 文件格式：.xls 或 .xlsx
- 表头结构：Type | Name | Address | DataType | InitValue | Comment | Attribute
- 数据从第3行开始（第2行为列名）

## 输出JSON格式
```json
[
  {
    "Name": "皮带电机报警",
    "Address": "MX2500.0"
  },
  {
    "Name": "玻璃阻挡气缸伸出",
    "Address": "MX2500.1"
  }
]
```

## 打包说明（开发者）
1. 确保已安装Python 3.7+
2. 双击运行 `打包成exe.bat`
3. 生成的exe文件在 `dist` 文件夹中

## 依赖库
- pandas: 读取Excel文件
- openpyxl: 支持.xlsx格式
- xlrd: 支持.xls格式
- pyinstaller: 打包工具
