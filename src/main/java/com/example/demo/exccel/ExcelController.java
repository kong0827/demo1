package com.example.demo.exccel;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.util.ArrayList;
import java.util.List;

@RestController
@RequestMapping("/excel")
public class ExcelController {

    @GetMapping("/parse")
    public void parseExcel() {
        String filePath = "path_to_your_excel_file.xlsx";
        List<BasicInfo> basicInfos = new ArrayList<>();
        List<FriendInfo> friendInfos = new ArrayList<>();

        EasyExcel.read(filePath).registerReadListener(new CustomListener(basicInfos, friendInfos)).sheet().doRead();
    }
}

class CustomListener extends AnalysisEventListener<List<String>> {
    private List<BasicInfo> basicInfos;
    private List<FriendInfo> friendInfos;

    private boolean isBasicInfo = false;
    private boolean isAgeRow = false;

    public CustomListener(List<BasicInfo> basicInfos, List<FriendInfo> friendInfos) {
        this.basicInfos = basicInfos;
        this.friendInfos = friendInfos;
    }

    @Override
    public void invoke(List<String> data, AnalysisContext context) {
        if (data.isEmpty()) {
            return; // 空行
        }
        if (!isBasicInfo) {
            // 处理基本信息行
            BasicInfo basicInfo = new BasicInfo();
            basicInfo.setName(data.get(0));
            basicInfos.add(basicInfo);
            isBasicInfo = true;
        } else {
            if (!isAgeRow) {
                // 处理可能是年龄的行
                if (data.get(0).toLowerCase().equals("age")) {
                    isAgeRow = true;
                    // 如果年龄在当前行，直接处理
                    if (data.size() > 1 && data.get(1) != null && data.get(1).matches("\\d+")) {
                        basicInfos.get(0).setAge(Integer.parseInt(data.get(1)));
                    }
                }
            } else {
                // 处理朋友信息行
                FriendInfo friendInfo = new FriendInfo();
                friendInfo.setFriendName(data.get(0));
                if (data.size() > 1 && data.get(1) != null && data.get(1).matches("\\d+")) {
                    friendInfo.setFriendAge(Integer.parseInt(data.get(1)));
                }
                if (data.size() > 2) {
                    friendInfo.setRemark(data.get(2));
                }
                friendInfos.add(friendInfo);
            }
        }
    }

    @Override
    public void doAfterAllAnalysed(AnalysisContext context) {
        // 处理完毕，可以对解析结果进行后续操作
        basicInfos.forEach(System.out::println);
        friendInfos.forEach(System.out::println);
    }
}
