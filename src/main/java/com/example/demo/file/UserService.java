package com.example.demo.file;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.write.builder.ExcelWriterBuilder;
import com.alibaba.excel.write.metadata.WriteSheet;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.ByteArrayResource;
import org.springframework.core.io.InputStreamSource;
import org.springframework.mail.javamail.JavaMailSender;
import org.springframework.mail.javamail.MimeMessageHelper;
import org.springframework.stereotype.Service;

import javax.mail.MessagingException;
import javax.mail.internet.MimeMessage;
import java.io.ByteArrayOutputStream;
import java.util.List;

@Service
public class UserService {

    @Autowired
    private UserRepository userRepository;

    @Autowired
    private JavaMailSender mailSender;

    public void exportUsersToExcelAndSendEmail(String toEmail) throws MessagingException {
        List<User> users = userRepository.findAll();
        List<SummaryData> summaryDataList = generateSummaryData(users);
        List<DetailData> detailDataList = generateDetailData(users);
        
        // Create Excel file
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriterBuilder excelWriterBuilder = EasyExcel.write(out);
        
        // Write summary data
        WriteSheet summarySheet = EasyExcel.writerSheet("Summary").head(SummaryData.class).build();
        excelWriterBuilder.sheet().doWrite(summaryDataList, summarySheet);
        
        // Write detail data
        WriteSheet detailSheet = EasyExcel.writerSheet("Detail").head(DetailData.class).build();
        excelWriterBuilder.sheet().doWrite(detailDataList, detailSheet);
        
        // Prepare email
        MimeMessage message = mailSender.createMimeMessage();
        MimeMessageHelper helper = new MimeMessageHelper(message, true);
        
        helper.setTo(toEmail);
        helper.setSubject("User Data");
        helper.setText("Please find the attached user data.");

        // Attach Excel file
        InputStreamSource attachment = new ByteArrayResource(out.toByteArray());
        helper.addAttachment("users.xlsx", attachment);

        // Send email
        mailSender.send(message);
    }

    private List<SummaryData> generateSummaryData(List<User> users) {
        // Generate summary data from users
        // Example logic:
        // Calculate totals based on group or other criteria
        // return summary data list
    }

    private List<DetailData> generateDetailData(List<User> users) {
        // Generate detail data from users
        // Example logic:
        // Convert users to detail data objects
        // return detail data list
    }
}
