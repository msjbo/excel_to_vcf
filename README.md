# Excel联系人转VCF工具

一个简单的Python应用程序，用于将Excel表格中的联系人信息转换为VCF格式，方便导入到手机或其他设备的通讯录中。

# 直接下载使用
  window应用直接下载目录中的excel_to_vcf.exe文件
## 功能特点

- 导入Excel文件（.xlsx/.xls）并在表格中预览内容
- 自动智能匹配Excel表格中的列与VCF联系人字段
- 可手动选择Excel列与VCF字段的对应关系
- 导出标准VCF格式文件，可直接导入到大多数设备的通讯录中

## 支持的VCF字段

- 姓名 (FN)
- 手机号码 (TEL)
- 电子邮件 (EMAIL)
- 地址 (ADR)
- 组织/公司 (ORG)
- 职位/头衔 (TITLE)

## 使用方法

1. 运行程序 `python excel_to_vcf.py`
2. 点击"导入Excel"按钮，选择包含联系人信息的Excel文件
3. 在表格中预览Excel内容
4. 检查并调整字段映射关系（程序会尝试自动匹配）
5. 点击"导出VCF"按钮，选择保存位置
6. 将生成的VCF文件导入到您的设备通讯录中

## 要求

- Python 3.6+
- 依赖库：pandas, tkinter (通常包含在Python标准库中)

## 安装依赖

```
pip install pandas
pip install openpyxl
```

## 注意事项

- Excel文件应包含联系人信息，列名最好能与联系人字段相关，以便自动匹配
- 建议Excel文件格式规范，每行代表一个联系人
- 程序会自动处理空值和无效数据 
