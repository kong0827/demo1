package com.example.demo.file;

import com.alibaba.excel.EasyExcel;
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
public class UserService1 {

    @Autowired
    private UserRepository userRepository;

    @Autowired
    private JavaMailSender mailSender;

    public void exportUsersToExcelAndSendEmail(String toEmail) throws MessagingException {
        List<User> users = userRepository.findAll();
        
        // Create Excel file
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        EasyExcel.write(out, User.class).sheet("Users").doWrite(users);
        
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
}
