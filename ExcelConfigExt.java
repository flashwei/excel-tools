package com.excel.utils.excel.refactor;

import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;

/**
 * @Author: Vachel Wang
 * @Date: 2017/5/3下午7:42
 * @Version: V1.0
 * @Description: 设值font必须要在设值cellStyle之前
 */
public class ExcelConfigExt {
    // 新创建sheet名称
    private String newSheetName;
    // 字体样式
    private XSSFFont font;
    // 单元格样式
    private XSSFCellStyle cellStyle;

    public String getNewSheetName() {
        return newSheetName;
    }

    public void setNewSheetName(String newSheetName) {
        this.newSheetName = newSheetName;
    }

    public XSSFFont getFont() {
        return font;
    }

    /**
     * 设值font必须要在设值cellStyle之前
     *
     * @param font 字体
     */
    public void setFont(XSSFFont font) {
        this.font = font;
    }

    public XSSFCellStyle getCellStyle() {
        return cellStyle;
    }

    public void setCellStyle(XSSFCellStyle cellStyle) {
        this.cellStyle = cellStyle;
        this.cellStyle.setFont(font);
    }
}
