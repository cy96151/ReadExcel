
# Excel通用读取工具

    Excel导入|大数据读取|不同版本

**Excel通用读取工具**是一个基于POI事件模式驱动的读取工具。通过工具类封装Excel文件对应的InputStream对象，并指定好回调接口，便可在回调接口中获取Excel文件中的每一行数据：
 
- **兼容不同版本** ：支持Excel常用的两个版本：Excel2003（.xls）和Excel2007（.xlsx）；
- **无须担心大数据** ：采用事件模式进行数据读取，告别传统用户模式构造数据对象后再进行读取，避免OutOfMemoryError；
- **其他扩展功能** ：支持读取过程中的跳过和终止，支持文件读取后保存（返回OutputStream），提供日期转换工具等。

-------------------

* [使用方法](#使用方法)
  * [示例](#示例)
  * [回调接口](#回调接口)
  * [数据基类](#数据基类)
  * [读取终止](#读取终止)
  * [线程安全](#线程安全)
* [注意事项](#注意事项)


## 使用方法
　　工具的使用十分简单，只需定义好回调接口，然后将需读取的文件流初始化，触发读取操作即可。下面一个简单的Demo来演示这个操作
### 示例
``` java
package cn.cy.test;

import java.io.FileInputStream;
import java.util.List;
import cn.cy.readexcel.read.ReadExcelBase;
import cn.cy.readexcel.rollback.ReadExcelRollBack;

public class Test implements ReadExcelRollBack {

	public static void main(String[] args) throws Exception {
		FileInputStream inputStream = new FileInputStream("E:\\test.xls");
		ReadExcelBase base = ReadExcelBase.create(inputStream, new Test());
		base.process();
	}

	@Override
	public void optRows(List<String> rowlist, int curRow, String sheetName, ReadExcelBase base) throws Exception {
		// 每读取一行数据，将会触发此方法。rowlist中存放此行数据的所有值。数据列标与所在List下标一致
		System.out.println(rowlist);
	}

	@Override
	public boolean judgeBreakSheet(String sheetName, ReadExcelBase base) {
		// 在读取每个Sheet页前会调用此方法，若无须读取当前Sheet页，可返回true
		return false;
	}
}
```

　　Test类实现了ReadExcelRollBack接口，当创建读取时，须将该对象的实例传入，以供回调。

### 回调接口

* 数据行处理（optRows）

``` java
/**
 * 数据行操作
 *
 * @param rowlist 当前读取的数据行
 * @param curRow 当前行号
 * @param sheetName 当前Sheet页名称
 * @param base 读取基类
 * @throws Exception 读取过程中产生的异常
 */
void optRows(List<String> rowlist, int curRow, String sheetName, ReadExcelBase base) throws Exception;
```
　　当每读取一行数据时，均会调用此方法。每行数据将会已String的形式保存在`rowlist`对象中。同时附带当前读取的相关属性，也可通过`数据基类base`获取其他当前数据。

　　在处理`rowlist`时，有以下注意事项：

　　1. `rowlist`数据下标与列号一致，若无数据则以`null`填充位置。`rowlist`的长度取决于最后一列有效值。

　　2. Excel中的数值单元格若精度过高，读取到的文本将会使用科学计数法，会造成精度丢失（Excel本身也有这个问题，大概能保存15-16位长度的数值）

　　3. 日期单元格读取后将会取到一个浮点数，其中整数部分表示1900年后到给定日期的天数，小数部分表示时分秒。`rowlist`中存放的是浮点数文本，可采用`org.apache.poi.ss.usermodel.DateUtil.getJavaDate`方法进行转换获取Date对象。

　　4. 在处理`rowlist`时，建议对空行进行特殊处理


* 当前Sheet是否读取（judgeBreakSheet）

``` java
/**
 * 判断读取当前sheet页内容是否跳过
 *
 * @param sheetName 当前Sheet页名称
 * @param base 读取基类
 * @return 若跳过此Sheet页的读取，返回true；否则返回false
 */
boolean judgeBreakSheet(String sheetName, ReadExcelBase base);
```
每读取一个Sheet页前，将会调用此方法，可根据sheetName和中的其他属性进行判断。

　　若返回true，则跳过此Sheet页的读取。

　　若返回false，则正常读取此Sheet页的内容

### 数据基类

当执行此行代码时，将会创建一个数据基类 ReadExcelBase

``` java
ReadExcelBase base = ReadExcelBase.create(inputStream, new Test());
```
　　在文件读取中，相关读取信息会保存在基类对象属性中。在回调方法里可通过`base.getXXX`的方式获取相关属性。

　　也可自行添加需要的属性，以供在回调方法中使用。

　　目前读取过程中能获取的属性如下：

属性 | 方法 | 说明 
- | :-: | -: 
sheetName | (String) base.getSheetName()| 获取当前sheet页名称
sheetIndex | (int) base.getSheetIndex() | 获取当前sheet页下标
curRow |(int) base.getCurRow()| 获取当前行号

### 读取终止

　　由于采用事件驱动读取文件，在默认情况下，只有文件内容全部读取完成，读取逻辑才会终止。若需在读取过程中手动终止，需在 `optRows` 方法中手动抛出异常。

　　目前已封装了两个异常类：

　　1. SheetContinueException

　　　　若在读取过程中抛出此异常，则终止当前Sheet页信息的读取

　　　　常用于读取指定行数据

　　2. SheetBreakException

　　　　若在读取过程中抛出此异常，则终止所有Sheet页信息的读取

　　　　常用于手动停止读取数据

### 线程安全

　　使用本工具时，针对每个文件流每次都会创建一个对象，无须担心线程安全问题。

　　但使用回调对象时，若使用单例对象且存在类属性时，需考虑回调实例的线程安全问题。

## 注意事项

1. 使用Excel创建的两种格式的文件均可以读取，但使用了`POI SXSSF`导出的Excel文件无法读取，原因是生成的ooxml内容格式不一致，需自行特殊处理。

2. 对Sheet页中的空行，本工具大部分情况下都是直接略过的。但空行中若存在空白值单元格（如空格），本行数据仍会被读取。建议在`optRows`方法中再进行一次校验。

3. 若有其他改进或补充，欢迎讨论。