package cn.cy.read;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;

import org.apache.poi.hssf.eventusermodel.AbortableHSSFListener;
import org.apache.poi.hssf.eventusermodel.HSSFEventFactory;
import org.apache.poi.hssf.eventusermodel.HSSFRequest;
import org.apache.poi.hssf.eventusermodel.HSSFUserException;
import org.apache.poi.hssf.record.BOFRecord;
import org.apache.poi.hssf.record.BlankRecord;
import org.apache.poi.hssf.record.BoolErrRecord;
import org.apache.poi.hssf.record.BoundSheetRecord;
import org.apache.poi.hssf.record.EOFRecord;
import org.apache.poi.hssf.record.FormulaRecord;
import org.apache.poi.hssf.record.LabelRecord;
import org.apache.poi.hssf.record.LabelSSTRecord;
import org.apache.poi.hssf.record.NumberRecord;
import org.apache.poi.hssf.record.Record;
import org.apache.poi.hssf.record.SSTRecord;
import org.apache.poi.hssf.record.StringRecord;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import cn.cy.exception.SheetBreakException;
import cn.cy.exception.SheetContinueException;
import cn.cy.rollback.ReadExcelRollBack;

/**
 * POI事件模式读取数据 支持版本：excel2003(.xls)
 *
 * @author cy96151
 */
public class ReadExcelOfHxls extends ReadExcelBase {
	/**
	 * 文件对象
	 */
	private POIFSFileSystem fs;

	/**
	 * 已读取的最后一行下标
	 * <p>
	 * 用于与当前单元格行下标进行判断是否为新的一行
	 * </p>
	 */
	private int lastRowNumber;

	/**
	 * 指向SSTRecord
	 * <p>
	 * 读取单元格字符串时通过此对象获取实际值
	 * </p>
	 */
	private SSTRecord sstRecord;

	/**
	 * 存放BoundSheetRecord对象
	 * <p>
	 * 对象顺序与Excel文件中的sheet页的实际顺序一致
	 * </p>
	 */
	private BoundSheetRecord[] orderedBSRs;
	private ArrayList<BoundSheetRecord> boundSheetRecords = new ArrayList<BoundSheetRecord>();

	/**
	 * 根据sheetIndex获取当前正在读取的sheet页属性
	 */
	private BoundSheetRecord thisSheetRecord;

	/**
	 * 单元格行下标
	 * <p>
	 * 用于公式缓存StringRecord的处理
	 * </p>
	 */
	private int nextRow;
	/**
	 * 单元格列下标
	 * <p>
	 * 用于公式缓存StringRecord的处理
	 * </p>
	 */
	private int nextColumn;
	/**
	 * 是否为StringRecord标志，若为true，则进入StringRecord逻辑块读取数据
	 * <p>
	 * 用于公式缓存StringRecord的处理
	 * </p>
	 */
	private boolean outputNextStringRecord;

	/**
	 * HSSFListener实现类
	 * <p>
	 * 重写了processRecord方法实现监听逻辑
	 * <p/>
	 */
	private ListenerImpl listener = null;

	public ReadExcelOfHxls(InputStream file, ReadExcelRollBack instance) throws Exception {
		this.fs = new POIFSFileSystem(file);
		this.listener = new ListenerImpl();
		this.instance = instance;
	}

	@Override
	public void process() throws Exception {
		HSSFEventFactory factory = new HSSFEventFactory();
		HSSFRequest request = new HSSFRequest();
		// 对所有类型的Record都设置此监听器
		request.addListenerForAllRecords(this.listener);
		factory.abortableProcessWorkbookEvents(request, fs);
	}

	private class ListenerImpl extends AbortableHSSFListener {
		/**
		 * 当前单元格行下标
		 */
		private int thisRow = -1;
		/**
		 * 当前单元格列下标
		 */
		private int thisColumn = -1;
		/**
		 * 当前单元格实际值
		 */
		private String thisStr = null;

		/**
		 * Record监听器
		 * <p>
		 * 读取每个有效Record的真实值，并拼凑出数据列表rowlist，当读取完一行后，调用rowDataProcess处理该行数据
		 * </p>
		 * <p>
		 * 目前监听的Record类型有：
		 * 
		 * <pre>
		 * BoundSheetRecord.sid
		 * BOFRecord.sid
		 * SSTRecord.sid
		 * BlankRecord.sid
		 * BoolErrRecord.sid
		 * FormulaRecord.sid
		 * StringRecord.sid
		 * LabelRecord.sid
		 * LabelSSTRecord.sid
		 * NumberRecord.sid
		 * EOFRecord.sid
		 * </pre>
		 * </p>
		 *
		 * @param record
		 *            Record对象，根据sid进行不同的处理
		 * @return 当前默认返回0
		 * @throws HSSFUserException
		 *             导入中产生的错误信息会封装为HSSFUserException，抛出后监听中止，导入结束
		 */
		@Override
		public short abortableProcessRecord(Record record) throws HSSFUserException {
			// 当前行下标
			thisRow = -1;
			// 当前列下标
			thisColumn = -1;
			// 当前值
			thisStr = null;
			// 返回值，若返回值非0，则终止读取
			short userCode = 0;
			try {
				// 根据Record的sid进行不同的处理
				switch (record.getSid()) {
				// sheet页信息
				case BoundSheetRecord.sid:
					processBoundSheetRecord(record);
					break;
				// sheet页开头
				case BOFRecord.sid:
					processBOFRecord(record);
					break;
				// 文本索引信息
				case SSTRecord.sid:
					processSSTRecord(record);
					break;
				// 空单元格
				case BlankRecord.sid:
					processBlankRecord(record);
					break;
				// bool或错误单元格
				case BoolErrRecord.sid:
					processBoolErrRecord(record);
					break;
				// 公式单元格
				case FormulaRecord.sid:
					processFormulaRecord(record);
					break;
				// 文本公式缓存
				case StringRecord.sid:
					processStringRecord(record);
					break;
				// 只读单元格
				case LabelRecord.sid:
					processLabelRecord(record);
					break;
				// 引用SSTRecord中的字符串
				case LabelSSTRecord.sid:
					processLabelSSTRecord(record);
					break;
				// 数值，日期单元格
				case NumberRecord.sid:
					processNumberRecord(record);
					break;
				// sheet页结尾
				case EOFRecord.sid:
					userCode = processEOFRecord(record);
					break;
				default:
					break;
				}

				// 遇到新行的操作
				if (thisRow != -1 && thisRow != lastRowNumber && rowArray.size() > 0) {
					userCode = rowDataProcess();
				}
				// 当前单元格的值非空时，将该值插入数据列表中
				if (thisStr != null) {
					rowListAddCell(thisColumn, thisStr);
				}
				// 更新当前行和当前列的值
				if (thisRow > -1) {
					lastRowNumber = thisRow;
					curRow = thisRow + 1;
				}
			} catch (Exception e) {
				throw new HSSFUserException(e.getMessage(), e);
			}
			return userCode;
		}

		/**
		 * Sheet页信息处理
		 *
		 * @param record
		 *            record
		 */
		private void processBoundSheetRecord(Record record) {
			boundSheetRecords.add((BoundSheetRecord) record);
		}

		/**
		 * Sheet页开始读取
		 * 
		 * <pre>
		 *     在此获取Sheet页名称，以及判断是否跳过读取
		 * </pre>
		 *
		 * @param record
		 *            record
		 */
		private void processBOFRecord(Record record) {
			BOFRecord br = (BOFRecord) record;
			if (br.getType() == BOFRecord.TYPE_WORKSHEET) {
				// Sheet页下标+1
				sheetIndex++;
				if (orderedBSRs == null) {
					// 将读取到的BoundSheetRecord对象根据Sheet页顺序进行排序
					orderedBSRs = BoundSheetRecord.orderByBofPosition(boundSheetRecords);
				}
				// 根据下标获取当前Shee页对应的BoundSheetRecord对象
				thisSheetRecord = orderedBSRs[sheetIndex];
				// 获取Sheet页名称
				sheetName = thisSheetRecord.getSheetname();
				// 判断当前sheet页是否跳过
				if (thisSheetRecord.isHidden() || thisSheetRecord.isVeryHidden()) {
					// 隐藏sheet页一律跳过
					breakSheet = true;
				} else {
					// 回调判断逻辑，判断Sheet页是否跳过读取
					breakSheet = instance.judgeBreakSheet(sheetName, ReadExcelOfHxls.this);
				}
				lastRowNumber = -1;
			}
		}

		/**
		 * 处理文本索引信息
		 *
		 * @param record
		 *            record
		 */
		private void processSSTRecord(Record record) throws Exception {
			sstRecord = (SSTRecord) record;
		}

		/**
		 * 处理空单元格
		 *
		 * @param record
		 *            record
		 */
		private void processBlankRecord(Record record) throws Exception {
			BlankRecord brec = (BlankRecord) record;
			thisRow = brec.getRow();
			thisColumn = brec.getColumn();
			// 将当前值置为空字符串
			thisStr = "";
		}

		/**
		 * 处理布尔和错误单元格
		 *
		 * @param record
		 *            record
		 */
		private void processBoolErrRecord(Record record) throws Exception {
			BoolErrRecord berec = (BoolErrRecord) record;
			thisRow = berec.getRow();
			thisColumn = berec.getColumn();
			if (berec.isBoolean()) {
				// 若单元格值为boolean，则获取布尔值
				thisStr = Boolean.toString(berec.getBooleanValue());
			} else {
				// 否则直接将值置空
				thisStr = "";
			}
		}

		/**
		 * 处理公式单元格
		 * 
		 * <pre>
		 *     当读取到公式单元格时，若公式值为数值，则直接获取
		 *     若公式值非数值，则会在下一个StringRecord对象中存放实际值
		 * </pre>
		 *
		 * @param record
		 *            record
		 */
		private void processFormulaRecord(Record record) throws Exception {
			FormulaRecord frec = (FormulaRecord) record;
			thisRow = frec.getRow();
			thisColumn = frec.getColumn();
			// 判断公式值是否为数字
			if (Double.isNaN(frec.getValue())) {
				// Formula result is a string
				// This is stored in the next record
				outputNextStringRecord = true;
				// 存放此单元格所在行列下标
				nextRow = frec.getRow();
				nextColumn = frec.getColumn();
			} else {
				thisStr = Double.toString(frec.getValue());
			}
		}

		/**
		 * 在公式文本缓存中获取实际值
		 *
		 * @param record
		 *            record
		 */
		private void processStringRecord(Record record) throws Exception {
			// 当outputNextStringRecord为true，说明之前已读到过公式单元格
			if (outputNextStringRecord) {
				// String for formula
				StringRecord srec = (StringRecord) record;
				// 获取实际值
				thisStr = srec.getString();
				// 获取之前公式单元格存放的行列下标
				thisRow = nextRow;
				thisColumn = nextColumn;
				outputNextStringRecord = false;
			}
		}

		/**
		 * 处理只读单元格
		 *
		 * @param record
		 *            record
		 */
		private void processLabelRecord(Record record) throws Exception {
			LabelRecord lrec = (LabelRecord) record;
			thisRow = lrec.getRow();
			thisColumn = lrec.getColumn();
			thisStr = lrec.getValue().trim();
		}

		/**
		 * 处理引用SSTRecord中的字符串
		 *
		 * @param record
		 *            record
		 */
		private void processLabelSSTRecord(Record record) throws Exception {
			LabelSSTRecord lsrec = (LabelSSTRecord) record;
			thisRow = lsrec.getRow();
			thisColumn = lsrec.getColumn();
			if (sstRecord == null) {
				// 若字符串索引为空，则将值置空
				thisStr = "";
			} else {
				// 根据索引获取字符串实际值
				thisStr = sstRecord.getString(lsrec.getSSTIndex()).toString().trim();
			}
		}

		/**
		 * 处理数值、日期单元格
		 *
		 * @param record
		 *            record
		 */
		private void processNumberRecord(Record record) throws Exception {
			NumberRecord numrec = (NumberRecord) record;
			thisRow = numrec.getRow();
			thisColumn = numrec.getColumn();
			// 将double值转换为String
			thisStr = Double.toString(numrec.getValue());
		}

		/**
		 * 处理sheet页结尾
		 *
		 * @param record
		 *            record
		 * @return 操作值，若返回0则继续读取，否则将终止读取
		 */
		private short processEOFRecord(Record record) throws Exception {
			EOFRecord er = (EOFRecord) record;
			if (rowArray.size() > 0) {
				lastRowNumber = -1;
				// 若当前行数据中存在值，需调用行数据处理方法
				return rowDataProcess();
			}
			return 0;
		}

		/**
		 * 行数据处理
		 *
		 * @return 操作值，若返回0则继续读取，否则将终止读取
		 * @throws Exception
		 *             读取过程中产生的非指定异常，抛出后将终止读取
		 */
		private short rowDataProcess() throws Exception {
			// 若当前sheet页非隐藏且没有被跳过读取，则回调业务逻辑处理接口
			if (!(thisSheetRecord.isHidden() || thisSheetRecord.isVeryHidden() || breakSheet)) {
				try {
					// 传入当前行的相关进行和自身基类，base中存放相关配置信息
					instance.optRows(rowArray, curRow, sheetName, ReadExcelOfHxls.this);
				} catch (SheetContinueException e) {
					// 解析过程中若抛出的异常信息为跳过本Sheet页读取，则将breakSheet改为true
					breakSheet = true;
				} catch (SheetBreakException e) {
					// 解析过程中若抛出的异常信息为终止整个文件的读取，则返回1，Excel将停止读取
					return 1;
				} catch (Exception e) {
					throw e;
				}
			}
			// 读取完成后将数据列表清空
			rowArray.clear();
			return 0;
		}
	}

	@Override
	public void saveFile(OutputStream stream) throws IOException {
		// 将文件对象保存至OutputStream中
		fs.writeFilesystem(stream);
	}
}
