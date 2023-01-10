1.以下划线开头的文件夹,文件名,和表名都会被忽略
2.表格第一行为参数名称,第二行为数值类型,第三行开始为配置数据
3.数值类型支持int,string,double,bool,Array
4.如果要转换成字典样式的json需要在excel表里创建数据透视表,否则不会转换成字典格式
5.如果excel只有一个sheet需要转换,则json数据里不包含sheet名称一层,否则会添加sheet层,比如:
demo.xlsx只有个sheet名为sheet1,转成json之后就是demo.json文件,内容为{.....},
如果除里sheet1之外还有sheet2,则转换之后文件内容为:{"sheet1":{...},"sheet2":{...}}

