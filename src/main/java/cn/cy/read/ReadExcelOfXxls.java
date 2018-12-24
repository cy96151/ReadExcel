package cn.cy.read;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTSheet;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STSheetState;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.WorkbookDocument;
import org.xml.sax.Attributes;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;
import org.xml.sax.helpers.XMLReaderFactory;

import cn.cy.exception.SheetBreakException;
import cn.cy.exception.SheetContinueException;
import cn.cy.rollback.ReadExcelRollBack;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;

/**
 * POI事件模式读取数据：excel2007(.xlsx) XSSF and SAX (Event API)
 * 
 * <pre>
 *     Excel2007及以上版本(.xlsx)的数据将使用xml进行数据存储，将一个excel文件后缀改为.zip，解压后即可看到文件存储结构
 *     本方法使用XSSFReader进行文件读取，该方法将提供各个Sheet页对应的xml文件流，再使用XMLReader对xml解析数据从而获取值
 * </pre>
 *
 * @author cy96151
 */
public class ReadExcelOfXxls extends ReadExcelBase {
	/**
	 * 字符串共享数据集
	 * <p>
	 * 对应XML文件：xl/sharedStrings.xml
	 * </p>
	 */
	private SharedStringsTable sst;

	/**
	 * 当前读取到的字符串
	 */
	private String lastContents;

	/**
	 * 读取XML时，标记下一个元素是否为SST的索引
	 */
	private boolean nextIsString;

	/**
	 * 当前列下标
	 */
	private int curCol = -1;

	private XSSFReader r;
	private OPCPackage pkg;

	public ReadExcelOfXxls(InputStream file, ReadExcelRollBack instance) throws Exception {
		pkg = OPCPackage.open(file);
		this.r = new XSSFReader(pkg);
		this.instance = instance;
	}

	@Override
	public void process() throws Exception {
		// 获取字符串共享表对象
		this.sst = r.getSharedStringsTable();
		// 获取xml解析对象
		XMLReader parser = fetchSheetParser();
		// 通过WorkbookDocument获取各Sheet页的CTSheet对象，可获取Sheet页名称和对应xml文件IO流
		List<CTSheet> sheetList = WorkbookDocument.Factory.parse(r.getWorkbookData()).getWorkbook().getSheets().getSheetList();
		for (CTSheet ctSheet : sheetList) {
			// Sheet页下标+1
			sheetIndex++;
			// 获取Sheet页名称
			sheetName = ctSheet.getName();
			// 隐藏sheet排除
			if (ctSheet.getState() == STSheetState.HIDDEN || ctSheet.getState() == STSheetState.VERY_HIDDEN) {
				continue;
			}
			// 判断当前sheet页是否跳过
			breakSheet = instance.judgeBreakSheet(sheetName, this);
			if (breakSheet) {
				continue;
			}
			// 获取Sheet页信息xml文件，文件路径：xl/worksheets/
			InputStream sheet = r.getSheet(ctSheet.getId());
			InputSource sheetSource = new InputSource(sheet);
			try {
				// 进行数据解析
				parser.parse(sheetSource);
			} catch (SheetContinueException e) {
				// 解析过程中若抛出的异常信息为跳过本Sheet页读取，则跳出本次循环
				continue;
			} catch (SheetBreakException e) {
				// 解析过程中若抛出的异常信息为终止整个文件的读取，则终止整个循环，Excel将停止读取
				break;
			} catch (SAXException e) {
				throw e;
			} finally {
				sheet.close();
			}
		}
	}

	/**
	 * 创建xml解析对象
	 * 
	 * <pre>
	 *     创建指定的ContentHandler并添加至此，数据读取时将调用指定的处理方法
	 * </pre>
	 *
	 * @return 创建的XMLReader对象
	 * @throws SAXException
	 *             创建对象失败将抛出异常
	 */
	private XMLReader fetchSheetParser() throws SAXException {
		XMLReader parser = XMLReaderFactory.createXMLReader("org.apache.xerces.parsers.SAXParser");
		parser.setContentHandler(new Handler());
		return parser;
	}

	/**
	 * XML数据处理类
	 */
	private class Handler extends DefaultHandler {
		@Override
		public void startElement(String uri, String localName, String name, Attributes attributes) throws SAXException {
			// 当元素名称为"c"时，则此处存放了一个单元格的值
			if ("c".equals(name)) {
				// 获取元素属性t的值
				String cellType = attributes.getValue("t");
				// 若属性t的值为"s"，则说明此单元格的实际值存放于字符串共享数据集中，子元素存放的时sst索引值，需将将nextIsString标记为true
				if (cellType != null && "s".equals(cellType)) {
					nextIsString = true;
				} else {
					nextIsString = false;
				}
				// 获取通过属性r的值获取单元格所在列下标
				curCol = cellIndexConver(attributes.getValue("r"));
			} else if ("row".equals(name)) {
				// 若元素名称为"row"，则是读取到了新的一行，根据属性r的值获取当前行下标
				String rowIndex = attributes.getValue("r");
				curRow = Integer.parseInt(rowIndex);
				// 清空行数据
				rowArray.clear();
				curCol = -1;
			}
			// 单元格值置空
			lastContents = "";
		}

		@Override
		public void endElement(String uri, String localName, String name) throws SAXException {
			// 当已读取完成的是v元素时，说明一个单元格读取完毕
			if ("v".equals(name)) {
				if (nextIsString) {
					// 若nextIsString为true，则lastContents值存放的是SST的索引值，需转换为实际值
					int idx = Integer.parseInt(lastContents);
					// 根据SST的索引值的到获取到单元格存储的实际字符串
					lastContents = new XSSFRichTextString(sst.getEntryAt(idx)).toString();
				}
				// 将单元格实际值加入rowlist中，并去掉字符串前后的空白符
				rowListAddCell(curCol, lastContents.trim());
			}

			// 如果元素名称为 row ，这说明已到行尾，调用 optRows() 方法
			if ("row".equals(name)) {
				try {
					instance.optRows(rowArray, curRow, sheetName, ReadExcelOfXxls.this);
				} catch (Exception e) {
					throw new SAXException(e.getMessage(), e);
				}
			}
		}

		@Override
		public void characters(char[] ch, int start, int length) throws SAXException {
			// 获取元素字符串
			lastContents += new String(ch, start, length);
		}

		/**
		 * 根据单元格元素的r属性转换为列下标
		 * 
		 * <pre>
		 *     r属性存放的数据规则为：列号+行号
		 *     excel列号是以大写英文字母的26进制的方式存放，需转换为10进制
		 * </pre>
		 *
		 * @param r
		 * @return
		 */
		private int cellIndexConver(String r) {
			int index = 0;
			// 通过将行号替换的方式获取列号值
			String letter = StringUtils.replace(r, String.valueOf(curRow), "");
			// 遍历列号值的每一位，将26进制转换为10进制
			for (int i = 0, length = letter.length(); i < length; i++) {
				int value = letter.charAt(letter.length() - 1 - i) - 64;
				index += value * Math.pow(26, i);
			}
			// 行号-1为数据数组下标
			return index - 1;
		}
	}

	@Override
	public void saveFile(OutputStream stream) throws IOException {
		// 将文件对象保存至OutputStream中
		pkg.save(stream);
	}
}
