# Excel2Json

- --excel_path excel 文件夹位置
- --json_path 输出 json 位置
- --script_path 输出脚本位置
- --script_template_path 脚本模板位置

- example 里是案例

- 第一行是字段名，第二行是类型，第三行是注释

- 一般的 bool ，int，float，string 都支持

- array_int 为 List<int> 其他基本类型类似 输入示例（1|2|3）

- dic_int_string 为 Dictionary<int,string> 其他基本类型类似 输入示例（1:a|2:b|3:c）

- //END 为结束行，如只想生成前三行，则可在第四行添加//END

- '#' 开头的行会忽略
