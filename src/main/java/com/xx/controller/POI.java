package com.xx.controller;

import com.alibaba.fastjson.JSON;
import com.xx.entity.User;
import com.xx.utils.ExcelUtils;
import com.xx.utils.ExcelUtils1;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.util.*;

/**
 * @author aqi
 * DateTime: 2020/5/27 2:02 下午
 * Description: poi导出
 *              HSSFWorkBook,
 *              XSSFworkbook,
 *              SXSSFworkbook
 */
@RestController
public class POI {

    /**
     * 控制数据的大小
     */
    private static Integer count = 1000000;
    private static List<User> userList = new ArrayList<>();
    private static ArrayList<Map<String, Object>> mapsList = new ArrayList<>();

    /**
     * 初始化数据
     */
    static {
        for (int i = 0; i < count; i++) {
//            userList.add(new User(i, "张三" + i, UUID.randomUUID().toString().replaceAll("-", ""), i % 2, i % 2 == 0, "这个数据是用户的备注信息，我想把这个数据弄长一点，现在这个长度感觉还不太行，应该要再长一点，现在这个长度我感觉差不多了，就这样吧，但是有时候导出的时候会导出很长的字段，不知道对导出的效率影响大不大，现在这个长度应该是够了", new Date(), "这里存放了一些其他字段", "", "其他2个就不存值了", ""));

            HashMap<String, Object> map = new HashMap<>(11);
            map.put("id", i);
            map.put("name", "张三" + i);
            map.put("password", UUID.randomUUID().toString().replaceAll("-", ""));
            map.put("gender", i % 2);
            map.put("live", i % 2 == 0);
            map.put("remarks", "这个数据是用户的备注信息，我想把这个数据弄长一点，现在这个长度感觉还不太行，应该要再长一点，现在这个长度我感觉差不多了，就这样吧，但是有时候导出的时候会导出很长的字段，不知道对导出的效率影响大不大，现在这个长度应该是够了");
            map.put("createTime", new Date());
            map.put("other", "这里存放了一些其他字段");
            map.put("msg1", "");
            map.put("msg2", "其他2个就不存值了");
            map.put("msg3", "");
            mapsList.add(map);

        }
    }

    @GetMapping("/getExcel")
    public void getExcel(HttpServletResponse response) throws IOException {
        String[] keys = {"序号", "姓名", "密码", "性别", "是否激活", "备注信息", "创建时间", "其他", "备注字段1", "备用字段2", "备用字段3"};
        String[] columnNames = {"id", "name", "password", "gender", "live", "remarks", "createTime", "other", "msg1", "msg2", "msg3"};
        String fileName = "demo.xlsx";

//        这个数据格式化在数据量很大的情况下,会导致堆内存溢出
//        String s = JSON.toJSONString(userList);
//        List<Map<String, Object>> maps = ExcelUtils.toList(s);


        long s = System.currentTimeMillis();
        ExcelUtils1.exportExcel(response, keys, columnNames, fileName, mapsList);
        long e = System.currentTimeMillis();
        System.out.println("导出Excel消耗的时间：" + (e - s));
    }

}
