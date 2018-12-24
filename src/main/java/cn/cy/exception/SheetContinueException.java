package cn.cy.exception;

import org.xml.sax.SAXException;

/**
 * 结束当前Sheet页读取异常类
 * 
 * <pre>
 * 若在读取过程中抛出此异常，则终止当前Sheet页信息的读取
 * 常用于读取指定行数据
 * </pre>
 * 
 * @author cy96151
 */
public class SheetContinueException extends SAXException {
	private static final long serialVersionUID = 1L;

	public SheetContinueException() {
		super();
	}
}
