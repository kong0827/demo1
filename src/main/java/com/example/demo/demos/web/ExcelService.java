package com.example.demo.demos.web;

import com.alibaba.excel.EasyExcel;
import org.springframework.stereotype.Service;

import java.io.File;
import java.io.IOException;
import java.util.List;

@Service
public class ExcelService {
    public File writeDataToExcel(List<Object> persons) throws IOException {
        String filePath = "persons.xlsx";
        EasyExcel.write(filePath, Object.class).sheet("Persons").doWrite(persons);
        return new File(filePath);
    }
}
