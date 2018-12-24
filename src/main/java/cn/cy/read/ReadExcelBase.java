package cn.cy.read;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.PushbackInputStream;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.apache.poi.POIXMLDocument;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import cn.cy.rollback.ReadExcelRollBack;

/**
 * Excel数据读取基类
 * 
 * <pre>
 *     基于POI事件驱动读取，针对03版(.xls)和07(.xlsx)版进行不同的处理
 *     在导入过程中，数据读取的相关值也存放于此
 * </pre>
 *
 * @author cy96151
 */
public abstract class ReadExcelBase {

	/*
	 * =========================================================================
	 * ===
	 */
	/* 通用Excel读取数据 */
	/*
	 * =========================================================================
	 * ===
	 */
	/**
	 * 当前sheet页名称
	 */
	protected String sheetName;

	/**
	 * 当前sheet页下标
	 */
	protected int sheetIndex = -1;

	/**
	 * 当前行数据
	 */
	protected List<String> rowArray = new ArrayList<String>();

	/**
	 * 当前行号
	 */
	protected int curRow = 0;

	/**
	 * 当前sheet页是否跳过
	 */
	protected boolean breakSheet = false;

	/**
	 * 回调实例
	 */
	protected ReadExcelRollBack instance;

	public String getSheetName() {
		return sheetName;
	}

	public int getCurRow() {
		return curRow;
	}

	public int getSheetIndex() {
		return sheetIndex;
	}

	public void setBreakSheet(boolean breakSheet) {
		this.breakSheet = breakSheet;
	}

	/**
	 * 通过判断文件版本，创建不同的读取事件驱动类 若文件版本是Excel03，则创建ReadExcelOfHxls
	 * 若文件版本是Excel07，则创建ReadExcelOfXxls 若无法识别，则抛出IllegalArgumentException异常
	 *
	 * @param inp
	 *            文件流
	 * @param instance
	 *            数据回调类
	 * @return ReadExcelBase 数据基类
	 * @throws Exception
	 *             若文件无法识别，则抛出IllegalArgumentException异常
	 */
	public static ReadExcelBase create(InputStream inp, ReadExcelRollBack instance) throws Exception {
		if (!inp.markSupported()) {
			inp = new PushbackInputStream(inp, 8);
		}

		if (POIFSFileSystem.hasPOIFSHeader(inp)) {
			return new ReadExcelOfHxls(inp, instance);
		}
		if (POIXMLDocument.hasOOXMLHeader(inp)) {
			return new ReadExcelOfXxls(inp, instance);
		}
		throw new IllegalArgumentException("无法识别Excel版本,请检查文件是否正常!");
	}

	/**
	 * 将任何符合日期格式的字符串转化为日期类型
	 * <p>
	 * parseStringToDate("2010-12-12") = Sun Dec 12 00:00:00 CST 2010
	 * parseStringToDate("20101212") = Sun Dec 12 00:00:00 CST 2010
	 * parseStringToDate("2010/12/12") = Sun Dec 12 00:00:00 CST 2010
	 * parseStringToDate("2010年12月12日") = Sun Dec 12 00:00:00 CST 2010
	 * parseStringToDate("2010 8 12") = Thu Aug 12 00:00:00 CST 2010
	 * parseStringToDate("20100802") = Mon Aug 02 00:00:00 CST 2010
	 * parseStringToDate("2010 8 2") = Mon Aug 02 00:00:00 CST 2010
	 * parseStringToDate("2010年8月2日") = Mon Aug 02 00:00:00 CST 2010
	 * <p>
	 * parseStringToDate("2010-12-12 05:04:03") = Sun Dec 12 05:04:03 CST 2010
	 * parseStringToDate("2010/12/12 05:04:03") = Sun Dec 12 05:04:03 CST 2010
	 * parseStringToDate("20101212 05:04:03") = Sun Dec 12 05:04:03 CST 2010
	 *
	 * @param sdate
	 *            要转换成日期的字符串
	 * @return Date
	 * @throws ParseException
	 *             格式转换异常
	 */
	public static Date parseStringToDate(String sdate) throws ParseException {
		String parse;
		parse = sdate.replaceFirst("^[0-9]{4}([^0-9]?)", "yyyy$1");
		parse = parse.replaceFirst("^[0-9]{2}([^0-9]?)", "yy$1");
		parse = parse.replaceFirst("([^0-9y]?)[0-9]{1,2}([^0-9]?)", "$1MM$2");
		parse = parse.replaceFirst("([^0-9M]?)[0-9]{1,2}( ?)", "$1dd$2");
		parse = parse.replaceFirst("( )[0-9]{1,2}([^0-9]?)", "$1HH$2");
		parse = parse.replaceFirst("([^0-9]?)[0-9]{1,2}([^0-9]?)", "$1mm$2");
		parse = parse.replaceFirst("([^0-9]?)[0-9]{1,2}([^0-9]?)", "$1ss$2");
		DateFormat format = new SimpleDateFormat(parse);
		return format.parse(sdate);
	}

	/**
	 * 触发文件读取驱动
	 *
	 * @throws Exception
	 *             文件读取中的异常抛出，包括校验异常、数据保存异常和手动抛出异常。 若抛出异常则终止数据导入，并抛出错误信息。
	 */
	public abstract void process() throws Exception;

	/**
	 * 文件读取完成后进行保存
	 *
	 * @param stream
	 *            保存文件的IO流对象，将会把文件内容写入到此IO流中
	 * @throws IOException
	 *             文件保存异常
	 */
	public abstract void saveFile(OutputStream stream) throws IOException;

	/**
	 * 数据行补空
	 * 
	 * <pre>
	 *     为保持数据行与表头行一致的效果，每次获取数据插入数据行列表时，要插入指定的下标位置，并将之前的空下标补空
	 * </pre>
	 *
	 * @param colIndex
	 *            数据下标
	 * @param cell
	 *            数据值
	 */
	protected void rowListAddCell(int colIndex, String cell) {
		for (int size = rowArray.size(), length = colIndex + 1; size < length; size++) {
			rowArray.add(null);
		}
		rowArray.set(colIndex, cell);
	}
}
