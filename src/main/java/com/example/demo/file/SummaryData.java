package com.example.demo.file;

import com.alibaba.excel.annotation.ExcelProperty;

public class SummaryData {
    @ExcelProperty("组别")
    private String group;
    
    @ExcelProperty("合计1")
    private int total1;
    
    @ExcelProperty("合计2")
    private int total2;
    
    // Getters and setters
}

import com.alibaba.excel.annotation.ExcelProperty;

public class DetailData {
    @ExcelProperty("明细标题1")
    private String detailTitle1;
    
    @ExcelProperty("明细标题2")
    private String detailTitle2;
    
    // Getters and setters
}
