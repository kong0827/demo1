package com.example.demo.demos.web;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.read.listener.PageReadListener;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.File;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@RestController
@RequestMapping("/excel")
public class ExcelController {

    @GetMapping("/parse")
    public void parseExcel() {
        String filePath = "path_to_your_excel_file.xlsx";
        List<BasicInfo> basicInfos = new ArrayList<>();
        List<FriendInfo> friendInfos = new ArrayList<>();

        EasyExcel.read(filePath, new PageReadListener<Map<Integer, String>>(dataList -> {
            for (int i = 0; i < dataList.size(); i++) {
                Map<Integer, String> data = dataList.get(i);
                if (i == 0) {
                    // 第一行，处理基本信息
                    BasicInfo basicInfo = new BasicInfo();
                    basicInfo.setName(data.get(1));
                    basicInfos.add(basicInfo);
                } else if (i == 1) {
                    // 第二行，处理基本信息
                    basicInfos.get(0).setAge(Integer.valueOf(data.get(1)));
                } else {
                    // 动态起始行，处理朋友信息
                    if (data.containsValue("朋友名称") || data.containsValue("朋友年龄") || data.containsValue("住址")) {
                        // 朋友信息标题行，跳过
                        continue;
                    }
                    FriendInfo friendInfo = new FriendInfo();
                    for (Map.Entry<Integer, String> entry : data.entrySet()) {
                        switch (entry.getValue()) {
                            case "朋友名称":
                                friendInfo.setFriendName(data.get(entry.getKey() + 1));
                                break;
                            case "朋友年龄":
                                friendInfo.setFriendAge(Integer.valueOf(data.get(entry.getKey() + 1)));
                                break;
                            case "住址":
                                friendInfo.setAddress(data.get(entry.getKey() + 1));
                                break;
                            default:
                                break;
                        }
                    }
                    friendInfos.add(friendInfo);
                }
            }
        })).sheet().doRead();

        // 打印或处理解析结果
        basicInfos.forEach(System.out::println);
        friendInfos.forEach(System.out::println);
    }
}
