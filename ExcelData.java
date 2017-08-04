package com.excel.utils.excel.refactor;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * @version V1.0
 * @Author: Vachel Wang
 * @Date: 2017/5/2
 * @Time: 下午8:19
 * @Description: 列的内容
 */
public abstract class ExcelData {

    private Object data;
    private Integer rowIndex;
    private Integer colIndex;
    private Integer colWidth;
    private Integer rowHeight;
    private Integer sheetIndex;
    private ExportExcelCustom exportExcelCustom;
    private ExcelConfigExt excelConfigExt;

    private static final Logger LOG = LoggerFactory.getLogger(ExcelData.class);

    /**
     * @param exportExcelCustom 处理Excel对象
     * @param data              数据
     * @param colIndex          列坐标
     * @param rowIndex          行坐标
     * @param colWidth          列宽
     * @param rowHeight         行高
     * @param sheetIndex        sheet下标
     * @param excelConfigExt    扩展配置
     */
    public ExcelData(ExportExcelCustom exportExcelCustom, Object data, Integer colIndex, Integer rowIndex, Integer colWidth, Integer rowHeight, Integer sheetIndex, ExcelConfigExt excelConfigExt) {
        this.setExportExcelCustom(exportExcelCustom);
        this.setData(data);
        this.setRowIndex(rowIndex);
        this.setColIndex(colIndex);
        this.setColWidth(colWidth);
        this.setRowHeight(rowHeight);
        this.setSheetIndex(sheetIndex);
        if (excelConfigExt == null)
            excelConfigExt = new ExcelConfigExt();
        this.setExcelConfigExt(excelConfigExt);
    }

    public abstract void fillData();

    public Integer getSheetIndex() {
        return sheetIndex;
    }

    public void setSheetIndex(Integer sheetIndex) {
        this.sheetIndex = sheetIndex;
    }

    public Object getData() {
        return data;
    }

    public void setData(Object data) {
        this.data = data;
    }

    public Integer getRowIndex() {
        return rowIndex;
    }

    public void setRowIndex(Integer rowIndex) {
        this.rowIndex = rowIndex;
    }

    public Integer getColIndex() {
        return colIndex;
    }

    public void setColIndex(Integer colIndex) {
        this.colIndex = colIndex;
    }

    public Integer getColWidth() {
        return colWidth;
    }

    public void setColWidth(Integer colWidth) {
        this.colWidth = colWidth;
    }

    public Integer getRowHeight() {
        return rowHeight;
    }

    public void setRowHeight(Integer rowHeight) {
        this.rowHeight = rowHeight;
    }

    public ExportExcelCustom getExportExcelCustom() {
        return exportExcelCustom;
    }

    public void setExportExcelCustom(ExportExcelCustom exportExcelCustom) {
        this.exportExcelCustom = exportExcelCustom;
    }

    public ExcelConfigExt getExcelConfigExt() {
        return excelConfigExt;
    }

    public void setExcelConfigExt(ExcelConfigExt excelConfigExt) {
        this.excelConfigExt = excelConfigExt;
    }
}
