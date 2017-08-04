package com.excel.utils.excel.refactor;

/**
 * @Author: Vachel Wang
 * @Date: 2017/5/2
 * @Time: 下午8:28
 * @Version: V1.0
 * @Description: 标题内容
 */
public class ExcelTitleData extends ExcelData {
    /**
     * @param exportExcelCustom 处理excel对象
     * @param dataArr           数据
     * @param colIndex          列坐标
     * @param rowIndex          行坐标
     * @param colWidth          列宽
     * @param rowHeight         行高
     * @param sheetIndex        sheetIndex
     * @param excelConfigExt    扩展配置
     */
    public ExcelTitleData(ExportExcelCustom exportExcelCustom, String[] dataArr, Integer colIndex, Integer rowIndex, Integer colWidth, Integer rowHeight, Integer sheetIndex, ExcelConfigExt excelConfigExt) {
        super(exportExcelCustom, dataArr, colIndex, rowIndex, colWidth, rowHeight, sheetIndex, excelConfigExt);
    }

    @Override
    public void fillData() {
        this.getExportExcelCustom().fillExcelTitleData(this);
    }
}
