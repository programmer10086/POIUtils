package com.xx.utils;

import com.alibaba.fastjson.JSON;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.net.URLEncoder;
import java.util.List;
import java.util.Map;
import java.util.concurrent.atomic.AtomicInteger;

/**
 * @author aqi
 * DateTime: 2020/5/28 9:17 上午
 * Description:
 *      使用HSSF进行Excel导出(不推荐使用)
 *          1.HSSF操作的是2003版本之前的Excel,扩展名是xls
 *          2.不希望方法调用的时候传递那么多参数,想要修改的参数直接修改方法内的静态参数就好了
 *          3.过多的样式我就不diy了,实在是太多了,默认定义了一个我觉得还行的样式,可以直接使用
 *      HSSF的缺陷：
 *          当导出数据超过65536条就会报错,抛出这个异常,网上有很多解决方案,我比较推荐不使用这种方式导出  java.lang.IllegalArgumentException: Invalid row number (65536) outside allowable range (0..65535)
 *      HSSF的效率
 *          10次一组跑了10次：
 *              1w条数据基本上在0.5秒以内,文件大小在2.5MB左右,偶尔存在波动情况
 *              5w条数据基本上在3秒左右,文件大小在11MB左右
 *
 *
 */
public class ExcelUtils {
    /**
     * 表头字体大小
     */
    private static String headerFontSize = "13";
    /**
     * 表头字体样式
     */
    private static String headerFontName = FontStyle.MicrosoftYahei.name;
    /**
     * 数据字体大小
     */
    private static String otherFontSize = "10";
    /**
     * 数据字体样式
     */
    private static String otherFontName = FontStyle.MicrosoftYahei.name;
    /**
     * 单元格宽度
     */
    private static Integer width = 30;
    /**
     * sheet的名字
     */
    private static String sheetName = "sheetName";
    /**
     * 是否开启表头样式,默认为true,开启
     */
    private static Boolean isOpeanHeaderStyle = true;
    /**
     * ##############是否开始其他数据样式,默认为false,关闭(不建议开启,数据量大时影响性能)################
     */
    private static Boolean isOpeanOtherStyle = false;

    /**
     * @param keys        对象属性对应中文名
     * @param columnNames 对象的属性名
     * @param fileName    文件名
     * @param list        需要导出的json数据
     * @description 使用HSSFWorkBook导出数据, HSSF导出数据存在一些问题
     */
    public static void exportExcel(HttpServletResponse response, String[] keys, String[] columnNames, String fileName, List<Map<String, Object>> list) throws IOException {
        // 创建一个工作簿
        HSSFWorkbook wb = new HSSFWorkbook();
        // 创建一个sheet
        HSSFSheet sh = wb.createSheet(sheetName);
        // 创建Excel工作表第一行,设置表头信息
        HSSFRow row0 = sh.createRow(0);
        for (int i = 0; i < keys.length; i++) {
            // 设置单元格宽度
            sh.setColumnWidth(i, 256 * width + 184);
            HSSFCell cell = row0.createCell(i);
            cell.setCellValue(keys[i]);
            // 是否开启表头样式
            if (isOpeanHeaderStyle) {
                // 创建表头样式
                HSSFCellStyle headerStyle = setCellStyle(wb, headerFontSize, headerFontName, "header");
                cell.setCellStyle(headerStyle);
            }
        }

//        for (int i = 0; i < list.size(); i++) {
//            // 循环创建行
//            HSSFRow row = sh.createRow(i + 1);
//            // 给这行的每列写入数据
//            for (int j = 0; j < columnNames.length; j++) {
//                HSSFCell cell = row.createCell(j);
//                // 以这样的方式取值,过滤掉不需要的字段
//                String value = String.valueOf(list.get(i).get(columnNames[j]));
//                cell.setCellValue(value);
//                // 是否开始其他数据样式
//                if (isOpeanOtherStyle) {
//                    // 设置数据样式
//                    HSSFCellStyle otherStyle = setCellStyle(wb, otherFontSize, otherFontName, "other");
//                    cell.setCellStyle(otherStyle);
//                }
//            }
//        }


        AtomicInteger i = new AtomicInteger();
        list.forEach(e ->{
            // 循环创建行
            HSSFRow row = sh.createRow(i.get() + 1);
            i.getAndIncrement();
            for (int j = 0; j < columnNames.length; j++) {
                HSSFCell cell = row.createCell(j);
                // 以这样的方式取值,过滤掉不需要的字段
                String value = String.valueOf(e.get(columnNames[j]));
                cell.setCellValue(value);
                // 是否开始其他数据样式
                if (isOpeanOtherStyle) {
                    // 设置数据样式
                    HSSFCellStyle otherStyle = setCellStyle(wb, otherFontSize, otherFontName, "other");
                    cell.setCellStyle(otherStyle);
                }
            }

        });



        response.setHeader("Content-Disposition", "attachment;filename=" + URLEncoder.encode(fileName, "UTF-8"));
        // 这个操作也非常的耗时,暂时不知道和什么有关,应该该和文件的大小有关
        wb.write(response.getOutputStream());
    }

    /**
     * @param wb       工作簿
     * @param fontSize 字体大小
     * @param fontName 字体名称
     * @return 工作簿样式
     * @description 设置Excel样式
     */
    private static HSSFCellStyle setCellStyle(HSSFWorkbook wb, String fontSize, String fontName, String boo) {
        // 创建自定义样式类
        HSSFCellStyle style = wb.createCellStyle();
        // 创建自定义字体类
        HSSFFont font = wb.createFont();
        // 设置字体样式
        font.setFontName(fontName);
        // 设置字体大小
        font.setFontHeightInPoints(Short.parseShort(fontSize));
        // 我这个版本的POI没找到网上的HSSFCellStyle
        // 设置对齐方式
        style.setAlignment(HorizontalAlignment.CENTER);
        // 数据内容设置边框实在太丑,容易看瞎眼睛,我帮你们去掉了
        if ("header".equals(boo)) {
            // 设置边框
            style.setBorderBottom(BorderStyle.MEDIUM);
            style.setBorderLeft(BorderStyle.MEDIUM);
            style.setBorderRight(BorderStyle.MEDIUM);
            style.setBorderTop(BorderStyle.MEDIUM);
            // 表头字体加粗
            font.setBold(true);
        }
        style.setFont(font);
        return style;
    }

    /**
     * 格式化数据(我发现这个操作非常的消耗时间,尽量不要使用到这个数据转化,如果是List<Map>就直接传,如果是Json稍微改一下上面的工具类,List<Bean>好像没什么比较好的处理手段)
     *
     * @param s json数据
     * @return 装换成List集合的数据
     */
    public static List<Map<String, Object>> toList(String s) {
        List<Map<String, Object>> list = (List) JSON.parse(s);
        return list;
    }

    /**
     * 找了半天也没找到可以diy的类,我自己写个吧
     */
    enum FontStyle {
        // 微软雅黑
        MicrosoftYahei("微软雅黑"),
        // 宋体
        TimesNewRoman("宋体"),
        // 楷体
        Italics("楷体"),
        // 幼圆
        YoungRound("幼圆");

        private String name;

        FontStyle(String name) {
            this.name = name;
        }
    }
}
