package cn.cy.rollback;

import java.util.List;

import cn.cy.read.ReadExcelBase;

/**
 * Excel读取回调接口
 *
 * @author cy96151
 */
public interface ReadExcelRollBack {
	/**
	 * 数据行操作
	 *
	 * @param rowlist
	 *            当前读取的数据行
	 * @param curRow
	 *            当前行号
	 * @param sheetName
	 *            当前Sheet页名称
	 * @param base
	 *            读取基类
	 * @throws Exception
	 *             读取过程中产生的异常
	 */
	void optRows(List<String> rowlist, int curRow, String sheetName, ReadExcelBase base) throws Exception;

	/**
	 * 判断读取当前sheet页内容是否跳过
	 *
	 * @param sheetName
	 *            当前Sheet页名称
	 * @param base
	 *            读取基类
	 * @return 若跳过此Sheet页的读取，返回true；否则返回false
	 */
	boolean judgeBreakSheet(String sheetName, ReadExcelBase base);
}
