package com.excel.utils.excel.refactor;

import java.util.ArrayList;
import java.util.List;

/**
 * @Author: Vachel Wang
 * @Date: 2017/5/2
 * @Time: 下午8:30
 * @Version: V1.0
 * @Description: 多行内容
 */
public class ExcelRowMultiData extends ExcelData {

    /**
     * @param exportExcelCustom 处理excel对象
     * @param rowData           行内容
     * @param colIndex          列坐标
     * @param rowIndex          行坐标
     * @param colWidth          列宽
     * @param rowHeight         行高
     * @param sheetIndex        sheet下标
     * @param excelConfigExt    扩展配置
     */
    public ExcelRowMultiData(ExportExcelCustom exportExcelCustom, List<ArrayList<ExcelData>> rowData, Integer colIndex, Integer rowIndex, Integer colWidth, Integer rowHeight, Integer sheetIndex, ExcelConfigExt excelConfigExt) {
        super(exportExcelCustom, rowData, colIndex, rowIndex, colWidth, rowHeight, sheetIndex, excelConfigExt);
    }

    @Override
    public void fillData() {
        this.getExportExcelCustom().fillRowMultiData(this);
    }
}
