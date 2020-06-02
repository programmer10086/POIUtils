package com.xx.utils;

import com.alibaba.fastjson.JSON;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.net.URLEncoder;
import java.util.List;
import java.util.Map;

/**
 * @author aqi
 * DateTime: 2020/5/28 11:50 上午
 * Description:
 *      使用SXSSF进行Excel导出(推荐使用)
 *          1.SXSSF用于大数据量导出,扩展名是xlsx
 *          2.不希望方法调用的时候传递那么多参数,想要修改的参数直接修改方法内的静态参数就好了
 *          3.过多的样式我就不diy了,实在是太多了,默认定义了一个我觉得还行的样式,可以直接使用
 *      SXSSF的缺陷：
 *          当导出数据超过1048576条就会报错,抛出这个异常,因为每个Sheet最多只能存1048576条数据,这时候需要将数据写到新的Sheet中,工具类已优化  java.lang.IllegalArgumentException: Invalid row number (1048576) outside allowable range (0..1048575)
 *      SXSSF的效率
 *          10次一组跑了10次：
 *              1w条数据基本上在0.5秒以内,文件大小在700KB左右
 *              5w条数据基本上在1.7秒左右,文件大小在3.5MB左右
 *              100w条数据基本上在35秒左右,文件大小在70MB左右(虽然内存没有溢出,但是cpu资源吃的厉害)
 *              200W条数据基本上在85秒左右,文件大小在140MB左右(做了多sheet处理,如果项目经理要一次性导出还要在5s内的话,你就把他鲨了吧)
 *
 *
 */
public class ExcelUtils1 {
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
     * 每个sheet存放的数据量
     */
    private static Integer sheetLength = 1000000;
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
     * @description 使用SXSSFWorkBook导出数据
     */
    public static void exportExcel(HttpServletResponse response, String[] keys, String[] columnNames, String fileName, List<Map<String, Object>> list) throws IOException {
        // 创建一个工作簿,每写100条数据就刷新数据出缓存,避免内存溢出
        SXSSFWorkbook wb = new SXSSFWorkbook(100);
        // 传入数据的大小
        int listSize = list.size();
        // 创建一个sheet
        SXSSFSheet sh = wb.createSheet(sheetName);
        // 设置这个sheet表头信息和样式
        setHeaderStyle(sh, keys, wb);
        // 用于计数,每100w时重新开始创建行
        int temp = 0;
        // 用于创建不同的sheetName
        int sheetNameEnd = 0;
        // 这个二重循环不知道有没有优化的空间了
        for (int i = 0; i < listSize; i++, temp++) {
            // 每100w重新创建一个新的sheet
            if (i % sheetLength == 0 && i != 0) {
                sheetNameEnd++;
                // 创建新的sheet
                sh = wb.createSheet(sheetName + sheetNameEnd);
                // 新的sheet设置新的单元格宽度
                setHeaderStyle(sh, keys, wb);
                temp = 0;
            }
            // 循环创建行
            SXSSFRow row = sh.createRow(temp + 1);
            // 给这行的每列写入数据
            for (int j = 0; j < columnNames.length; j++) {
                SXSSFCell cell = row.createCell(j);
                // 以这样的方式取值,过滤掉不需要的字段
                String value = String.valueOf(list.get(i).get(columnNames[j]));
                cell.setCellValue(value);
                // 是否开始其他数据样式
                if (isOpeanOtherStyle) {
                    // 设置数据样式
                    CellStyle otherStyle = setCellStyle(wb, otherFontSize, otherFontName, "other");
                    cell.setCellStyle(otherStyle);
                }
            }
        }
        response.setHeader("Content-Disposition", "attachment;filename=" + URLEncoder.encode(fileName, "UTF-8"));
        // 这个操作也非常的耗时,应该该和文件的大小有关,后续看看能不能优化
        wb.write(response.getOutputStream());
    }

    /**
     * 设置表头样式,在大数据情况下每个sheet都要执行一次,所以抽出出来了
     * @param sh sheet
     * @param keys 对象属性对应中文名
     * @param wb 工作簿
     */
    private static void setHeaderStyle (SXSSFSheet sh, String[] keys, SXSSFWorkbook wb) {
        // 创建Excel工作表第一行,设置表头信息
        SXSSFRow row0 = sh.createRow(0);
        for (int i = 0; i < keys.length; i++) {
            // 设置单元格宽度
            sh.setColumnWidth(i, 256 * width + 184);
            SXSSFCell cell = row0.createCell(i);
            cell.setCellValue(keys[i]);
            // 是否开启表头样式
            if (isOpeanHeaderStyle) {
                // 创建表头样式
                CellStyle headerStyle = setCellStyle(wb, headerFontSize, headerFontName, "header");
                cell.setCellStyle(headerStyle);
            }
        }
    }

    /**
     * @param wb       工作簿
     * @param fontSize 字体大小
     * @param fontName 字体名称
     * @return 工作簿样式
     * @description 设置Excel样式
     */
    private static CellStyle setCellStyle(SXSSFWorkbook wb, String fontSize, String fontName, String boo) {
        // 创建自定义样式类
        CellStyle style = wb.createCellStyle();
        // 创建自定义字体类
        Font font = wb.createFont();
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
     * 数据量达到200w的时候堆内存直接就爆了,我错了我错了,数量达千万别用     java.lang.OutOfMemoryError: Java heap space
     *
     * @param s json数据
     * @return 转换成List集合的数据
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
