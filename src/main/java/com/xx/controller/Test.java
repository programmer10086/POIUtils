//package com.xx.controller;
//
//import com.xx.entity.User;
//import org.apache.poi.xssf.streaming.SXSSFCell;
//import org.apache.poi.xssf.streaming.SXSSFRow;
//import org.apache.poi.xssf.streaming.SXSSFSheet;
//import org.apache.poi.xssf.streaming.SXSSFWorkbook;
//import org.springframework.web.bind.annotation.GetMapping;
//import org.springframework.web.bind.annotation.RestController;
//
//import javax.servlet.http.HttpServletResponse;
//import java.io.IOException;
//import java.net.URLEncoder;
//import java.time.Duration;
//import java.time.LocalDateTime;
//import java.util.ArrayList;
//import java.util.List;
//
///**
// * @author aqi
// * DateTime: 2020/5/18 5:13 下午
// * Description: No Description
// */
//@RestController
//public class Test {
//
//    @GetMapping("/outExcel")
//    public void outPutExcel(HttpServletResponse response) throws IOException {
//        List<User> userList = new ArrayList<>();
//        int limit = 30000;
//        for (int i = 0; i < limit; i++) {
//            userList.add(new User(i, "name is " + i, "password is " + i, 1, "阿七测试用的辅助核算", "阿七辅助明细测试摘要", "100.00", "借", "2019-10-31", "供应商", "其他业务收入", ""));
//        }
//
//        long s = System.currentTimeMillis();
//
//        // 每次写入100条数据就刷新数据到缓存中
//        SXSSFWorkbook wb = new SXSSFWorkbook(100);
//        SXSSFSheet sh = wb.createSheet("sheet");
//
//        for (int i = 0; i < userList.size(); i++) {
//            SXSSFRow row = sh.createRow(i);
//            User user = userList.get(i);
//
//            SXSSFCell cell1 = row.createCell(0);
//            cell1.setCellValue(user.getId());
//
//            SXSSFCell cell2 = row.createCell(1);
//            cell2.setCellValue(user.getName());
//
//            SXSSFCell cell3 = row.createCell(2);
//            cell3.setCellValue(user.getPassword());
//
//            SXSSFCell cell4 = row.createCell(3);
//            cell4.setCellValue(user.getGender());
//
//            SXSSFCell cell5 = row.createCell(4);
//            cell5.setCellValue(user.getMessage());
//
//            SXSSFCell cell6 = row.createCell(5);
//            cell6.setCellValue(user.getMessage1());
//
//            SXSSFCell cell7 = row.createCell(6);
//            cell7.setCellValue(user.getMessage2());
//
//            SXSSFCell cell8 = row.createCell(7);
//            cell8.setCellValue(user.getMessage3());
//
//            SXSSFCell cell9 = row.createCell(7);
//            cell9.setCellValue(user.getMessage4());
//
//            SXSSFCell cell10 = row.createCell(8);
//            cell10.setCellValue(user.getMessage5());
//
//            SXSSFCell cell11 = row.createCell(9);
//            cell11.setCellValue(user.getMessage6());
//
//            SXSSFCell cell12 = row.createCell(10);
//            cell12.setCellValue(user.getMessage7());
//
//
//        }
//
//        String fileName = "demo.xlsx";
//        response.setHeader("Content-Disposition", "attachment;filename=" + URLEncoder.encode(fileName, "UTF-8"));
//
//        long s10 = System.currentTimeMillis();
//
//        wb.write(response.getOutputStream());
//
//
//
//        wb.close();
//
//        long e = System.currentTimeMillis();
//
//        System.out.println("进行write操作消耗的时间：" + (e - s10));
//        System.out.println("总共消耗的时间：" + (e - s));
//
//    }
//
//}
