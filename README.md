# combine_excel

## 小姐妹提的需求
每家公司都有一个表，表的格式相同（每个表的工作簿sheet名，每个sheet的格式都相同），要把所有公司的数据合到一张表里，每个sheet对应合并。用网上的方法经常出问题，一直没有一劳永逸的方法，还是手动一张一张粘贴。每个月都有20+张表要合并。

<img src='https://pic.downk.cc/item/5ea43695c2a9a83be5ec5654.png' width=450>

## 方案
用 pandas 读取 excel 文件作为 DataFrame，对 DataFrame 合并然后再导出为 excel。

pandas 用 `read_excel` 读取多个 sheet 的 excel 表：
```python
pd.read_excel(path, header=None, sheet_name=None)
```
- sheet_name=None 表示读取所有的 sheet，如果是单独读取某个 sheet，取值写 sheet 的名字就可以。
- 输出的格式是字典 {'sheet 01': pandas.DataFrame, 'sheet 02': pandas.DataFrame, 'sheet 03': pandas.DataFrame}，key 就是 sheet 名，value 就是那张表的 DataFrame 格式。

对于每张 sheet，遍历文件夹内所有 excel 表合成该张 sheet 的汇总。

存为 excel 文件用 `to_excel`：
```python
writer = pd.ExcelWriter(output_path)
sheet.to_excel(writer, index=False,encoding='utf-8',header=None,sheet_name=s)
```
- output_path 就是新生成的 excel 表的路径，不需要先建立 excel 文件，直接生成新的 excel 文件。
- 每次的 to_excel 存储 sheet_name 为 s 的数据，连着用几次 to_excel，不同的 sheet_name 就生成不同的 sheet。

考虑到生成文件的管理，或许有时候修改了原数据又要重新生成，或许同名会难以分清。在生成的文件名添加了生成时间作为区分。

使用注意事项，就是最好保证读取的文件夹内只有需要合并的 xlsx 文件，虽然加了正则表达式只识别后缀为 xlsx 的文件。

鉴于小姐妹的财务报表是比较规范的，每张表的格式是一样的，懒得考虑异常处理。最后打包成 exe 发送给小姐妹使用~
