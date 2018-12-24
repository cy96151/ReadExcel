package cn.cy.test;

import java.io.FileInputStream;
import java.util.List;

import cn.cy.read.ReadExcelBase;
import cn.cy.rollback.ReadExcelRollBack;

public class Test implements ReadExcelRollBack {

	public static void main(String[] args) throws Exception {
		FileInputStream inputStream = new FileInputStream("E:\\test.xls");
		ReadExcelBase base = ReadExcelBase.create(inputStream, new Test());
		base.process();
	}

	@Override
	public void optRows(List<String> rowlist, int curRow, String sheetName, ReadExcelBase base) throws Exception {
		System.out.println(rowlist);
		// 可在通过base获取当前读取的其他相关信息，如 当前sheet页下标 等
		// 在此手动抛出异常，可终止读取，如抛出 SheetBreakException 或 SheetContinueException
	}

	@Override
	public boolean judgeBreakSheet(String sheetName, ReadExcelBase base) {
		// 在读取每个Sheet页前会调用此方法，若无须读取当前Sheet页，可返回true
		return false;
	}

}
