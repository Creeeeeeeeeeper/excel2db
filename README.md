# excel2db
A small tool for converting excel, csv to SQLite databse.

使用`excel2db.exe -h` 可以查看使用方法：
```bash
将Excel和CSV文件转换为SQLite数据库
Usage: excel2db.exe [OPTIONS]
Options:
  -p, --path <FILE>  指定要转换的文件路径
  -h, --help         Print help
  -V, --version      Print version
```

***更详细的食用方法如下：***

## 食用方法

### 1.快速转换

使用cmd直接在程序后面添加文件路径，相当于把excel直接拖到excel2db.exe上，也就是构造一个这样的命令：`excel2db.exe D:/demo.xlsx`，如果是单工作表会直接输出同名db文件，如果是多工作表，只会转换第一个工作表，并且输出的db文件会命名成`原文件名 - 第一个表名.db`。

### 2.参数转换

使用cmd，直接运行`excel2db.exe`可以按照提示进行一步一步转换，注意这样的转换方式输入文件路径时不要带引号。

使用cmd，使用`-p`参数，例如`excel2db.exe -p D:/demo.xlsx`，跟上一步差不多，也是根据提示进行转换



最新更新 v1.1：

特性：**使用显式事务（conn.transaction()）包裹批量插入，避免每条数据单独提交导致的IO瓶颈。**
使用方法一快速转换一个十万行的excel大约需要3s

Bug：有时会数据重复转换，如果转换后的db文件大小明显大于excel文件，建议查看条数是否一致，可以构建sql语句进行查询：`SELECT COUNT(*) AS total_count FROM data;`

### 欢迎提出issue!

