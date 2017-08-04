package com.excel.utils.excel.refactor;

/**
 * @Author: Vachel Wang
 * @Date: 2017/5/2
 * @Time: 下午8:27
 * @Version: V1.0
 * @Description: 图片内容
 */
public class ExcelPictureData extends ExcelData {
    /**
     * @param exportExcelCustom 处理Excel对象
     * @param data              数据内容
     * @param startColIndex     开始列坐标
     * @param startRowIndex     开始行坐标
     * @param endColIndex       结束列坐标
     * @param endRowIndex       结束行坐标
     * @param colWidth          列宽
     * @param rowHeight         行高
     * @param sheetIndex        sheet下标
     * @param excelConfigExt    扩展配置
     */
    public ExcelPictureData(ExportExcelCustom exportExcelCustom, byte[] data, int startColIndex, int startRowIndex, int endColIndex, int endRowIndex, Integer colWidth, Integer rowHeight, Integer sheetIndex, ExcelConfigExt excelConfigExt) {
        super(exportExcelCustom, data, startColIndex, startRowIndex, colWidth, rowHeight, sheetIndex, excelConfigExt);
        this.startColIndex = startColIndex;
        this.startRowIndex = startRowIndex;
        this.endColIndex = endColIndex;
        this.endRowIndex = endRowIndex;
    }

    private Integer startRowIndex;
    private Integer endRowIndex;
    private Integer startColIndex;
    private Integer endColIndex;

    public Integer getStartRowIndex() {
        return startRowIndex;
    }

    public void setStartRowIndex(Integer startRowIndex) {
        this.startRowIndex = startRowIndex;
        this.setRowIndex(startRowIndex);
    }

    public Integer getEndRowIndex() {
        return endRowIndex;
    }

    public void setEndRowIndex(Integer endRowIndex) {
        this.endRowIndex = endRowIndex;
    }

    public Integer getStartColIndex() {
        return startColIndex;
    }

    public void setStartColIndex(Integer startColIndex) {
        this.startColIndex = startColIndex;
        this.setColIndex(startColIndex);
    }

    public Integer getEndColIndex() {
        return endColIndex;
    }

    public void setEndColIndex(Integer endColIndex) {
        this.endColIndex = endColIndex;
    }

    @Override
    public void fillData() {
        this.getExportExcelCustom().fillPictureData(this);
    }
}
