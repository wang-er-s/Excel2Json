# Excel2Json

- -e excel文件夹位置
- -j 输出json位置
- -c 输出脚本位置
- -t 脚本模板位置
- -b true为bson,false为json
- -l 为脚本里的json加载路径，如 "Assets/Res/Config/#json_name.bytes"，#json_name会替换成队名名字

- example里是案例，bat只有win可以使用，py win和mac都可以用

- 第一行是字段名，第二行是类型，第三行是注释

- 一般的bool ，int，float，string都支持

- array_int为List<int>  其他基本类型类似 输入示例（1|2|3）

- dic_int_string 为Dictionary<int,string> 其他基本类型类似  输入示例（1:a|2:b|3:c）

- //END为结束行，如只想生成前三行，则可在第四行添加//END

- '#' 开头的行会忽略