package cn.cy.exception;

import org.xml.sax.SAXException;

/**
 * 终止所有Sheet页读取异常类
 * 
 * <pre>
 * 若在读取过程中抛出此异常，则终止所有Sheet页信息的读取
 * 常用于手动停止读取数据
 * </pre>
 * 
 * @author cy96151
 */
public class SheetBreakException extends SAXException {

	private static final long serialVersionUID = 1L;

	public SheetBreakException() {
		super();
	}
}
