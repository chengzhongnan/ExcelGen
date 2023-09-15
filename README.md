<style>
table {
	width: 90%;
	background: #ccc;
	margin: 10px auto;
	border-collapse: collapse;
	/*border-collapse:collapse合并内外边距
(去除表格单元格默认的2个像素内外边距*/
}
th,
td {
	height: 25px;
	line-height: 25px;
	text-align: center;
	border: 1px solid #ccc;
}
th {
	background: #eee;
	font-weight: normal;
}
tr {
	background: #fff;
}
tr:hover {
	background: #cc0;
}
td a {
	color: #06f;
	text-decoration: none;
}
td a:hover {
	color: #06f;
	text-decoration: underline;
}
</style>

## 文件名与Sheet页签规范

1.  Excel文件名规范: `文件名描述配置文件功能.xlsm/xlsx`，例如 `物品表.xlsx`，`公会表.xlsx`
2.  一个Excel文件可以包含多个Sheet页签,Sheet页签必须为英文命名，并且格式为:`{Name}_Table`,例如:`Item_Table`, `Equaipment_Table`
3.  当Sheet页签以`Sheet`开头时,该页签的内容会被生成工具忽略，可以用来写说明等信息,例如:`Sheet1`，`Sheet_Item表_说明`，里面的内容不会被工具解析，所以可以随意编写
4.  当Sheet页签以`Config`结尾时,该页签的内容会被视枚举配置表,里面的内容会被工具解析并生成枚举代码,例如:`Item_Table_Config`, `Equaipment_Table_Config`
5.  当Sheet页签以`_Config`结尾时,该表只有第一列的数据会被解析，后面列的数据会被忽略，可以将第二列当做备注列使用，详细规则见下文
6.  第一行表头和第二行表头必须为该列数据的英文描述，并且采用大驼峰命名法，即单词首字母大写，其他小写，单词之间不采用下划线，而是直接合并到一起
	
## 配置文件格式规范

1. 配置文件头为四行结构，第一行为主层级，第二行为次层级，第三行为数据类型，第四行为数据描述，大多数情况下，表格为平面结构，第一行与第二行相同，最后编译出来为最简单的json数据，例如：
   
	<table>
    <tr>
        <td>ItemId</td>
        <td>PropName</td>
        <td>Describe</td>
        <td>PropType</td>
        <td>PropQualityType</td>
        <td>Duration</td>
    </tr>
    <tr>
        <td>ItemId</td>
        <td>PropName</td>
        <td>Describe</td>
        <td>PropType</td>
        <td>PropQualityType</td>
        <td>Duration</td>
    </tr>
    <tr>
        <td>int</td>
        <td>string</td>
        <td>string</td>
        <td>int</td>
        <td>int</td>
        <td>int</td>
    </tr>
    <tr>
        <td>具Id</td>
        <td>道具名称</td>
        <td>道具描述</td>
        <td>道具的分类</td>
        <td>品质</td>
        <td>物品有效时长(以秒为单位，0表示无限)</td>
    </tr>
    <tr>
        <td>1</td>
        <td>ItemTable_PropName_1</td>
        <td>ItemTable_Describe_1</td>
        <td>1</td>
        <td>4</td>
        <td>0</td>
    </tr>
</table>

<p align="left">该表格编译出来的数据为：</p>

```json
[{"ItemId":1,"PropName":"ItemTable_PropName_1","Describe":"ItemTable_Describe_1","PropType":1,"PropQualityType":4,"Duration":0}]
```
	
2. 明确有联系的数据可以合并到一起，方式为将第一行对应列合并单元格到一起，其他不变，例如：
   	<table>
		<tr>
			<td>Id</td>
			<td colspan="3">Book</td>
		</tr>
		<tr>
			<td>Id</td>
			<td>Name</td>
			<td>Autor</td>
			<td>PublishAt</td>
		</tr>
		<tr>
			<td>int</td>
			<td>string</td>
			<td>string</td>
			<td>string</td>
		</tr>
		<tr>
			<td>Id</td>
			<td>名称</td>
			<td>作者</td>
			<td>出版日期</td>
		</tr>
		<tr>
			<td>1</td>
			<td>西游记</td>
			<td>吴承恩</td>
			<td>明朝</td>
		</tr>
		</table>
 	编译出来的数据为：
```json
	[{"Id":1,"Book":{"Name":"西游记","Autor":"吴承恩","PublishAt":"明朝"}}]
```

3. 数组的支持，当需要填写的数据为一组数组（可以为字符串数组，整数数组，浮点数组），采用与2类似的方式，第一行为对应列合并单元格到一起，第二行为多个同样的标签重复，例如：
   <table>
    <tr>
        <td>Id</td>
        <td colspan="3">Score</td>
        <td>Average</td>
    </tr>
    <tr>
        <td>Id</td>
        <td>Score</td>
        <td>Score</td>
        <td>Score</td>
        <td>Average</td>
    </tr>
    <tr>
        <td>int</td>
        <td>int</td>
        <td>int</td>
        <td>int</td>
        <td>int</td>
    </tr>
    <tr>
        <td>Id</td>
        <td>分数</td>
        <td>分数</td>
        <td>分数</td>
        <td>平均分</td>
    </tr>
    <tr>
        <td>1</td>
        <td>85</td>
        <td>81</td>
        <td>76</td>
        <td>81</td>
    </tr>
</table>
上面的表中Score是一个数组，编译出来的数据为：

```json
[{"Id":1,"ScoreList":[85,81,76],"Average":81}]
```

4. 结构体数组的支持，和上述规则类似，当第二行是间隔重复的时候，表示的是一个结构体数组，结构体数组里面的标签必须有明确的循环节，否则会导致编译错误，例如：
   <table>
    <tr>
        <td>ID</td>
        <td>Prob</td>
        <td  colspan="6">Items</td>
    </tr>
    <tr>
        <td>ID</td>
        <td>Prob</td>
        <td>ItemId</td>
        <td>Num</td>
        <td>ItemId</td>
        <td>Num</td>
        <td>ItemId</td>
        <td>Num</td>
    </tr>
    <tr>
        <td>int</td>
        <td>int</td>
        <td>int</td>
        <td>int</td>
        <td>int</td>
        <td>int</td>
        <td>int</td>
        <td>int</td>
    </tr>
    <tr>
        <td>ID</td>
        <td>概率</td>
        <td>物品id</td>
        <td>物品数量</td>
        <td>物品id</td>
        <td>物品数量</td>
        <td>物品id</td>
        <td>物品数量</td>
    </tr>
    <tr>
        <td>1</td>
        <td>100</td>
        <td>1001</td>
        <td>1000</td>
        <td>1002</td>
        <td>1</td>
        <td>2001</td>
        <td>1</td>
    </tr>
    <tr>
        <td>2</td>
        <td>200</td>
        <td>1001</td>
        <td>1000</td>
        <td>1003</td>
        <td>1</td>
        <td>2002</td>
        <td>1</td>
    </tr>
    
</table>

上面的Items是一个{ItemId, Num}的结构体数组，编译出来的数据为：
```json
[
    {
        "ID": 1,
        "Prob": 100,
        "ItemsList": [
            {
                "ItemId": 1001,
                "Num": 1000
            },
            {
                "ItemId": 1002,
                "Num": 1
            },
            {
                "ItemId": 2001,
                "Num": 1
            }
        ]
    },
    {
        "ID": 2,
        "Prob": 200,
        "ItemsList": [
            {
                "ItemId": 1001,
                "Num": 1000
            },
            {
                "ItemId": 1003,
                "Num": 1
            },
            {
                "ItemId": 2002,
                "Num": 1
            }
        ]
    }
]
```
5.  Config枚举配置文件规则：以Config结尾的文件为配置文件配表，该表表头与其他表相同，不过表格格式相对比较固定，一般设定为3列，第一列是最重要的，会在代码中生成对应的枚举结构，但是不会生成对应的xml和json文件，第二列为枚举名称描述，第三列为枚举对应数值，是一串从0开始的顺序数值，注意：配置表不能在中间插入行，如果要增加配置，请加在当前文件的最后一行。
   建议把第一行数据配置为None
   例如：
   <table>
    <tr>
        <td>Index</td>
        <td>Desc</td>
        <td>Value</td>
    </tr>
    <tr>
        <td>Index</td>
        <td>Desc</td>
        <td>Value</td>
    </tr>
    <tr>
        <td>string</td>
        <td>string</td>
        <td>int</td>
    </tr>
    <tr>
        <td>索引名</td>
        <td>描述</td>
        <td>配置值</td>
    </tr>
	    <tr>
        <td>None</td>
        <td></td>
        <td>0</td>
    </tr>
    <tr>
        <td>RollExp</td>
        <td>每次投掷骰子获得经验</td>
        <td>1</td>
    </tr>
    <tr>
        <td>RollExpLimit</td>
        <td>每日投掷骰子获得经验上限</td>
        <td>2</td>
    </tr>
    <tr>
        <td>TaskCountMax</td>
        <td>任务获取的最大值</td>
        <td>3</td>
    </tr>
    <tr>
        <td>EventCountMax</td>
        <td>事件获取的最大值</td>
        <td>4</td>
    </tr>
    <tr>
        <td>MailSearchLimit</td>
        <td>一次性可查询邮件的最大数</td>
        <td>5</td>
    </tr>
    <tr>
        <td>PlayDiceLimit</td>
        <td>每日投掷骰子次数上限(目前按RollExp/..Limit来算 )</td>
        <td>6</td>
    </tr>
</table>

生成的枚举为：
```csharp
    public enum EnumSystemConfig
    {
		None,
        RollExp,
        RollExpLimit,
        TaskCountMax,
        EventCountMax,
        MailSearchLimit,
        PlayDiceLimit,
    }
