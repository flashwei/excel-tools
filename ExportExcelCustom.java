package com.excel.utils.excel.refactor;

import javax.servlet.http.HttpServletResponse;

import org.apache.poi.ss.usermodel.Workbook;

/**
 * @Author: Vachel Wang
 * @Date: 2017/5/2
 * @Time: 下午8:28
 * @Version: V1.0
 * @Description: 重构自定义导出
 * 1.具体导出规则由子类实现，方便后期维护和添加导出其它格式，符合开闭原则
 * 2.重构代码结构去掉冗余代码，优化导出方式
 */
public interface ExportExcelCustom {
    /**
     * 模板替换字符串
     */
    String REPLACE_CHAR = "%s";
    /**
     * 高度换算率=1px
     */
    double HEIGIT_TIMES = 15.625;
    /**
     * 宽度度换算率=1px
     */
    double WIDTH_TIMES = 35.7;

    /**
     * 导出
     *
     * @param fileName 文件名
     * @param response response
     */
    void export(String fileName, HttpServletResponse response);

    /**
     * 填充文本数据
     *
     * @param textData 数据
     */
    void fillTextData(ExcelTextData textData);

    /**
     * 填充图片数据
     *
     * @param pictureData 数据
     */
    void fillPictureData(ExcelPictureData pictureData);

    /**
     * 填充单行数据
     *
     * @param rowSignData 数据
     */
    void fillRowSignData(ExcelRowSignData rowSignData);

    /**
     * 填充多行数据
     *
     * @param rowMultiData 数据
     */
    void fillRowMultiData(ExcelRowMultiData rowMultiData);

    /**
     * 填充标题数据
     *
     * @param titleData 数据
     */
    void fillExcelTitleData(ExcelTitleData titleData);

    /**
     * 获取Workbook对象
     *
     * @return Workbook
     */
    Workbook getWorkbook();
}
