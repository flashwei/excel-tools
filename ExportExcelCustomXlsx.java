package com.excel.utils.excel.refactor;


import java.io.File;
import java.io.IOException;
import java.io.OutputStream;
import java.net.URLEncoder;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import javax.servlet.http.HttpServletResponse;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * 通过自定义模板导出excel 仅支持导出xlsx格式
 *
 * @author Vachel.Wang
 * @version V1.1
 * @date 2016年7月4日 下午1:52:23
 */
public class ExportExcelCustomXlsx implements ExportExcelCustom {

	private XSSFWorkbook workbook = null;
	private static final Logger LOG = LoggerFactory.getLogger(ExportExcelCustomXlsx.class);

	/**
	 * 构造
	 *
	 * @param templatePath
	 *            模板路径，可填项
	 * @throws IOException
	 * @throws InvalidFormatException
	 */
	public ExportExcelCustomXlsx(String templatePath) throws IOException, InvalidFormatException {
		if (templatePath == null) {
			workbook = new XSSFWorkbook();
		} else {
			File file = new File(templatePath);
			if (!file.exists())
				throw new IOException("文件不存在：" + templatePath);
			workbook = (XSSFWorkbook) WorkbookFactory.create(file);
		}
	}

	/**
	 * 导出
	 *
	 * @param fileName
	 * @param response
	 */
	@Override
	public void export(String fileName, HttpServletResponse response) {
		OutputStream outputStream = null;
		try {
			response.reset();
			response.setContentType("application/octet-stream; charset=utf-8");
			response.setHeader("Content-Disposition", "attachment; filename=" + URLEncoder.encode(fileName, "UTF-8"));
			outputStream = response.getOutputStream();
			workbook.write(outputStream);
		} catch (IOException e) {
			LOG.info(e.getMessage(), e);
			throw new GenericException("导出excel异常");
		} finally {
			try {
				if (null != outputStream) {
					outputStream.flush();
					outputStream.close();
				}
			} catch (IOException e) {
				LOG.info(e.getMessage(), e);
			}
		}
	}

	/**
	 * 根据下标获取sheet，如果不存在则创建
	 *
	 * @param sheetIndex
	 * @param newSheetName
	 * @return
	 */
	private XSSFSheet getSheet(Integer sheetIndex, String newSheetName) {
		XSSFSheet sheet;
		if (workbook.getNumberOfSheets() < (sheetIndex + 1)) {
			if (newSheetName == null)
				newSheetName = "newsheet" + (sheetIndex + 1);
			sheet = workbook.createSheet(newSheetName);
		} else {
			sheet = workbook.getSheetAt(sheetIndex);
		}
		return sheet;
	}

	/**
	 * 根据行下标获取Row不存在则创建
	 *
	 * @param sheet
	 * @param rowIndex
	 * @return
	 */
	private XSSFRow getRow(XSSFSheet sheet, Integer rowIndex) {
		XSSFRow row = sheet.getRow(rowIndex);
		if (row == null)
			row = sheet.createRow(rowIndex);
		return row;
	}

	/**
	 * 插入一行
	 *
	 * @param sheet
	 * @param rowIndex
	 * @return
	 */
	private XSSFRow createRow(XSSFSheet sheet, Integer rowIndex) {
		XSSFRow row = null;
		if (sheet.getRow(rowIndex) != null) {
			int lastRowNo = sheet.getLastRowNum();
			sheet.shiftRows(rowIndex, lastRowNo, 1);
		}
		row = sheet.createRow(rowIndex);
		return row;
	}

	/**
	 * 根据下标获取cell不存在则创建
	 *
	 * @param row
	 * @param colIndex
	 * @return
	 */
	private XSSFCell getCell(XSSFRow row, Integer colIndex) {
		XSSFCell cell = row.getCell(colIndex);
		if (cell == null)
			cell = row.createCell(colIndex);
		return cell;
	}

	/**
	 * 设置单元格宽高度
	 *
	 * @param excelData
	 * @param sheet
	 * @param row
	 */
	private void setWidthAndHeight(ExcelData excelData, XSSFSheet sheet, XSSFRow row) {
		if (excelData.getRowHeight() != null)
			row.setHeight((short) (HEIGIT_TIMES * excelData.getRowHeight()));
		if (excelData.getColWidth() != null)
			sheet.setColumnWidth(excelData.getColIndex(), (short) (WIDTH_TIMES * excelData.getColWidth()));
	}

	@Override
	public void fillTextData(ExcelTextData textData) {
		// 获取sheet
		XSSFSheet sheet = getSheet(textData.getSheetIndex(), textData.getExcelConfigExt().getNewSheetName());
		// 获取row
		XSSFRow row = getRow(sheet, textData.getRowIndex());
		// 获取cell
		XSSFCell cell = getCell(row, textData.getColIndex());
		// 设置尺寸
		setWidthAndHeight(textData, sheet, row);
		// 填充数据
		fillCellContent(cell, textData.getData(), textData.getExcelConfigExt());
	}

	@Override
	public void fillPictureData(ExcelPictureData pictureData) {
		// 获取sheet
		XSSFSheet sheet = getSheet(pictureData.getSheetIndex(), pictureData.getExcelConfigExt().getNewSheetName());
		// 获取row
		XSSFRow row = getRow(sheet, pictureData.getStartRowIndex());
		// 获取cell
		XSSFCell cell = getCell(row, pictureData.getStartColIndex());
		// 设置尺寸
		setWidthAndHeight(pictureData, sheet, row);
		// 填充数据
		fillCellPicture(sheet, pictureData);
	}

	@Override
	public void fillRowSignData(ExcelRowSignData rowSignData) {
		List<ExcelData> rowData = (List<ExcelData>) rowSignData.getData();
		XSSFSheet sheet = getSheet(rowSignData.getSheetIndex(), rowSignData.getExcelConfigExt().getNewSheetName());
		createRow(sheet, rowSignData.getRowIndex());
		for (int j = 0; j < rowData.size(); j++) {
			ExcelData excelData = rowData.get(j);

			// 图片
			if (excelData instanceof ExcelPictureData) {
				ExcelPictureData pictureData = (ExcelPictureData) excelData;
				pictureData.setStartColIndex(rowSignData.getColIndex() + j);
				pictureData.setEndColIndex(rowSignData.getColIndex() + j + 1);
				pictureData.setStartRowIndex(rowSignData.getRowIndex());
				pictureData.setEndRowIndex(rowSignData.getRowIndex() + 1);
				pictureData.setSheetIndex(rowSignData.getSheetIndex());

			} else if (excelData instanceof ExcelTextData) { // 文本
				ExcelTextData excelTextData = (ExcelTextData) excelData;
				excelTextData.setColIndex(rowSignData.getColIndex() + j);
				excelTextData.setRowIndex(rowSignData.getRowIndex());
				excelTextData.setSheetIndex(rowSignData.getSheetIndex());
			}
			excelData.fillData();
		}
	}

	@Override
	public void fillRowMultiData(ExcelRowMultiData rowMultiData) {
		List<ArrayList<ExcelData>> multiData = (List<ArrayList<ExcelData>>) rowMultiData.getData();
		// 行
		for (int k = 0; k < multiData.size(); k++) {
			ArrayList<ExcelData> columnDataList = multiData.get(k);
			ExcelRowSignData signData = new ExcelRowSignData(rowMultiData.getExportExcelCustom(), columnDataList,
					rowMultiData.getColIndex(), rowMultiData.getRowIndex() + k, rowMultiData.getColWidth(),
					rowMultiData.getRowHeight(), rowMultiData.getSheetIndex(), rowMultiData.getExcelConfigExt());
			signData.fillData();
		}
	}

	@Override
	public void fillExcelTitleData(ExcelTitleData titleData) {
		if (titleData.getExcelConfigExt().getFont() == null) {
			/* 默认标题字体 */
			XSSFFont titleFont = workbook.createFont();
			titleFont.setFontHeightInPoints((short) 16);
			titleFont.setFontName("Courier New");
			titleFont.setBold(true);
			titleData.getExcelConfigExt().setFont(titleFont);
		}
		if (titleData.getExcelConfigExt().getCellStyle() == null) {
			/* 默认标题样式 */
			XSSFCellStyle titleStyle = workbook.createCellStyle();
			titleStyle.setAlignment(CellStyle.ALIGN_CENTER);
			titleStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
			titleData.getExcelConfigExt().setCellStyle(titleStyle);
		}
		// 获取sheet
		XSSFSheet sheet = getSheet(titleData.getSheetIndex(), titleData.getExcelConfigExt().getNewSheetName());
		// 获取row
		XSSFRow row = getRow(sheet, titleData.getRowIndex());

		String[] titleDataArr = (String[]) titleData.getData();
		for (int i = 0; i < titleDataArr.length; i++) {
			XSSFCell cell = getCell(row, titleData.getColIndex() + i);
			cell.setCellValue(titleDataArr[i]);
			cell.setCellStyle(titleData.getExcelConfigExt().getCellStyle());
		}
	}

	/**
	 * 填充内容
	 *
	 * @param cell
	 * @param data
	 */
	private void fillCellContent(XSSFCell cell, Object data, ExcelConfigExt configExt) {
		if (cell == null)
			return;
		if (data == null) {
			String cellData = getStringCellValue(cell);
			cell.setCellValue(cellData.replaceAll(REPLACE_CHAR, ""));
			return;
		}
		String cellData = getStringCellValue(cell);
		// 字符串
		if (data instanceof String) {
			String strData = data.toString();
			if (cellData.indexOf(REPLACE_CHAR) != -1) {
				cellData = String.format(cellData, strData.split(";"));
			} else {
				cellData += strData;
			}
			cell.setCellValue(cellData);
		}
		// 整数
		else if (data instanceof Integer) {
			Integer integerData = (Integer) data;
			if (cellData.equals("")) {
				cell.setCellValue(integerData);
			} else if (cellData.indexOf(REPLACE_CHAR) != -1) {
				cellData = cellData.replaceAll(REPLACE_CHAR, String.valueOf(integerData));
				cell.setCellValue(cellData);
			} else {
				cell.setCellValue(integerData);
			}
		}
		// 长整数
		else if (data instanceof Long) {
			Long longData = (Long) data;
			if (cellData.equals("")) {
				cell.setCellValue(longData);
			} else if (cellData.indexOf(REPLACE_CHAR) != -1) {
				cellData = cellData.replaceAll(REPLACE_CHAR, String.valueOf(longData));
				cell.setCellValue(cellData);
			} else {
				cell.setCellValue(longData);
			}
		}
		// 小数
		else if (data instanceof Double) {
			Double doubleData = (Double) data;
			if (cellData.equals("")) {
				cell.setCellValue(doubleData);
			} else if (cellData.indexOf(REPLACE_CHAR) != -1) {
				cellData = cellData.replaceAll(REPLACE_CHAR, String.valueOf(doubleData));
				cell.setCellValue(cellData);
			} else {
				cell.setCellValue(doubleData);
			}
		}
		// 时间
		else if (data instanceof Date) {
			Date dateData = (Date) data;
			if (cellData.equals("")) {
				cell.setCellValue(dateData);
			} else if (cellData.indexOf(REPLACE_CHAR) != -1) {
				cellData = cellData.replace(REPLACE_CHAR, DateUtil.formatYYYYMMDDHHMMSS(dateData));
				cell.setCellValue(cellData);
			} else {
				cell.setCellValue(DateUtil.formatYYYYMMDDHHMMSS(dateData));
			}
		}
		// 数组
		else if (data instanceof String[]) {
			String[] strArr = (String[]) data;
			String str = "";

			if (cellData.equals("")) {
				for (String s : strArr) {
					str += "," + s;
				}
				if (str.length() > 0)
					str = str.substring(1);
			} else if (cellData.indexOf(REPLACE_CHAR) != -1) {
				str = String.format(cellData, strArr);
			}
			cell.setCellValue(str);
		}
		// 其它
		else {
			cell.setCellValue(data + "");
		}
		cell.setCellStyle(configExt.getCellStyle());
	}

	/**
	 * 填充图片
	 *
	 * @param sheet
	 * @param pictureData
	 */
	private void fillCellPicture(XSSFSheet sheet, ExcelPictureData pictureData) {
		if (pictureData == null || pictureData.getData() == null)
			return;
		byte bytes[] = (byte[]) pictureData.getData();
		int pictureIdx = workbook.addPicture(bytes, XSSFWorkbook.PICTURE_TYPE_JPEG);
		Drawing drawing = sheet.createDrawingPatriarch();
		XSSFClientAnchor anchor = new XSSFClientAnchor(0, 0, 255, 255, pictureData.getStartColIndex(),
				pictureData.getStartRowIndex(), pictureData.getEndColIndex(), pictureData.getEndRowIndex());
		drawing.createPicture(anchor, pictureIdx);
	}

	@Override
	public XSSFWorkbook getWorkbook() {
		return workbook;
	}

}
