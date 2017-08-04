package com.excel.utils.excel.refactor;

/**
 * @Author: Vachel Wang
 * @Date: 2017/5/2
 * @Time: 下午8:24
 * @Version: V1.0
 * @Description: 单列内容，内容暂时未做限制，仅支持
 */
public class ExcelTextData extends ExcelData {
    /**
     * @param exportExcelCustom 处理excel对象
     * @param data              数据
     * @param colIndex          列坐标
     * @param rowIndex          行坐标
     * @param colWidth          列宽
     * @param rowHeight         行高
     * @param sheetIndex        sheet下标
     * @param excelConfigExt    扩展配置
     */
    public ExcelTextData(ExportExcelCustom exportExcelCustom, Object data, Integer colIndex, Integer rowIndex, Integer colWidth, Integer rowHeight, Integer sheetIndex, ExcelConfigExt excelConfigExt) {
        super(exportExcelCustom, data, colIndex, rowIndex, colWidth, rowHeight, sheetIndex, excelConfigExt);
    }

    @Override
    public void fillData() {
        this.getExportExcelCustom().fillTextData(this);
    }

}
