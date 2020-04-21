package com.example.easyexceltest.controller;

import com.alibaba.excel.EasyExcel;
import com.example.easyexceltest.bean.DemoData;
import com.example.easyexceltest.bean.Student;
import com.example.easyexceltest.listener.DemoDataListener;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.net.URLEncoder;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.List;

/**
 * create by sunzhongwei on 2020/04/20
 * Desc:
 */
@RestController
public class TestController {
    //
    @GetMapping("/test1")
    public List<DemoData> test1(){
        // 写法1：
        String fileName = "/home/hanhai/szw/preview/easy-excel-test/test.xlsx";
        // 这里 需要指定读用哪个class去读，然后读取第一个sheet 文件流会自动关闭
        DemoDataListener demoDataListener = new DemoDataListener();
        EasyExcel.read(fileName, DemoData.class, demoDataListener).sheet().doRead();

        System.out.println(demoDataListener.getList());
        return demoDataListener.getList();
    }

    @GetMapping("/test2")
    public void test2(){
        // 写法1：
        String fileName = "/home/hanhai/szw/preview/easy-excel-test/testWrite.xlsx";
        // 这里 需要指定读用哪个class去读，然后读取第一个sheet 文件流会自动关闭
        EasyExcel.write(fileName, DemoData.class).sheet("模板").doWrite(data());
    }


    @GetMapping("/test3")
    public void test3(HttpServletResponse response) throws IOException {
        // 这里注意 有同学反应使用swagger 会导致各种问题，请直接用浏览器或者用postman
        response.setContentType("application/vnd.ms-excel");
        response.setCharacterEncoding("utf-8");
        // 这里URLEncoder.encode可以防止中文乱码 当然和easyexcel没有关系
        String fileName = URLEncoder.encode("测试", "UTF-8");
        response.setHeader("Content-disposition", "attachment;filename=" + fileName + ".xlsx");
        EasyExcel.write(response.getOutputStream(), DemoData.class).sheet("模板").doWrite(data());
    }


    @GetMapping("/test4")
    public void test4(HttpServletResponse response) throws IOException {
        // 这里注意 有同学反应使用swagger 会导致各种问题，请直接用浏览器或者用postman
        response.setContentType("application/vnd.ms-excel");
        response.setCharacterEncoding("utf-8");
        // 这里URLEncoder.encode可以防止中文乱码 当然和easyexcel没有关系

        Student stu1 = Student.builder()
                .id(1)
                .name("张三")
                .age(18)
                .build();

        Student stu2 = Student.builder()
                .id(2)
                .name("李四")
                .age(19)
                .build();

        Student stu3 = Student.builder()
                .id(3)
                .name("王五")
                .age(24)
                .build();

        List<Student> students = Arrays.asList(stu1, stu2, stu3);


        String fileName = URLEncoder.encode("学生", "UTF-8");
        response.setHeader("Content-disposition", "attachment;filename=" + fileName + ".xlsx");
        EasyExcel.write(response.getOutputStream(), Student.class).sheet("模板szw").doWrite(students);
    }


    private List<DemoData> data() {
        List<DemoData> list = new ArrayList<DemoData>();
        for (int i = 0; i < 10; i++) {
            DemoData data = new DemoData();
            data.setString("字符串" + i);
            data.setDate(new Date());
            data.setDoubleData(0.56);
            list.add(data);
        }
        return list;
    }
}
