# ClasstableToIcal
Convert Classtable to iCal using Pything and Excel as data source.

该工具可以方便地将课程表转换为 `.ics` 格式以导入各种设备的「日历」中。

> 修改自原项目, 目前只适合HITSZ自建教务系统导出的Excel课表文件

## Usage

> 目前暂时使用1.2.0版本, 以便支持xlsx文件

```shell
pip install uuid xlrd==1.2.0
```

然后执行 `main.py`：

```shell
python main.py
```

测试环境为：Python 3.9.0，Windows 10 x64.

## 文件中格式解释

export.xlsx为教务系统导出的Excel文件

## Feature

只支持HITSZ

## License

LGPLv3
