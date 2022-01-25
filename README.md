# Excel to cfg

### excel格式 

| 行数  | 内容     | 示例                              |
|-----|--------|---------------------------------|
| 1   | 中文描述   | 随意.给策划看                         |
| 2   | 代码中变量名 | id val...                       |
| 3   | 类型     | int/long/string/int[]/long[]    |
| 4   | 输出范围   | ALL/CLIENT/SERVER/IGNORE(忽略输出列) |



>  如果第二行变量名为 ".ignore". 并且某一行的该列标注为 true  yes 1 等值. 这行数据忽略.



### sheet 规则

1. sheet 如果为end . 后面的内容就不会读取.
2. sheet 名包含 `c.` 表示仅客户端需要. `s.` 表示仅服务器需要
3. sheet 名包含其它比如`lua.` 表示客户端需要按照lua输出. 会去 home目录下的`.dTools/ejs` 找对应的ejs模板



