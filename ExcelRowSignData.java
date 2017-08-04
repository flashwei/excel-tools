package com.excel.utils.excel.refactor;

import java.util.List;

/**
 * @Author: Vachel Wang
 * @Date: 2017/5/2
 * @Time: 下午8:29
 * @Version: V1.0
 * @Description: 单行内容
 */
public class ExcelRowSignData extends ExcelData {
    /**
     * @param exportExcelCustom 处理excel对象
     * @param rowData           数据内容
     * @param colIndex          填充列坐标
     * @param rowIndex          填充行坐标
     * @param colWidth          列宽
     * @param rowHeight         行高
     * @param sheetIndex        sheet下标
     * @param excelConfigExt    扩展配置
     */
    public ExcelRowSignData(ExportExcelCustom exportExcelCustom, List<ExcelData> rowData, Integer colIndex, Integer rowIndex, Integer colWidth, Integer rowHeight, Integer sheetIndex, ExcelConfigExt excelConfigExt) {
        super(exportExcelCustom, rowData, colIndex, rowIndex, colWidth, rowHeight, sheetIndex, excelConfigExt);
    }


    @Override
    public void fillData() {
        this.getExportExcelCustom().fillRowSignData(this);
    }
}
