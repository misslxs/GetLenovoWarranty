# GetLenovoWarranty
批量查询联想电脑保修时间

## 使用方法
从excel读取sn数据，输入输出列自行在warranty_2_excel 进行更改。
```
更改excel文件存储路径
path = '/Users/xxx/lenovo.xlsx'
```
输入:C列为序列号。输出列: N列为`保修开始时间`，O列`保修结束时间`，P列`保修状态`。
