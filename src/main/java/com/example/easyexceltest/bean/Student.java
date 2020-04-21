package com.example.easyexceltest.bean;

import com.alibaba.excel.annotation.ExcelProperty;
import lombok.Builder;
import lombok.Data;

/**
 * create by sunzhongwei on 2020/04/20
 * Desc:
 */
@Data
@Builder
public class Student {
    @ExcelProperty("学号")
    private Integer id;
    @ExcelProperty("姓名")
    private String name;
    @ExcelProperty("年龄")
    private Integer age;
}
