package com.example.demo.demos.web;

@RestController
@RequestMapping("/api")
public class PersonController {
    @Autowired
    private PersonService personService;

    @Autowired
    private ExcelService excelService;

    @Autowired
    private EmailService emailService;

    @GetMapping("/sendPersonData")
    public String sendPersonData() {
        try {
            // 1. 从数据库读取数据
            List<Person> persons = personService.getAllPersons();
            
            // 2. 将数据写入Excel文件
            File excelFile = excelService.writeDataToExcel(persons);

            // 3. 发送带Excel附件的邮件
            emailService.sendEmailWithAttachment("recipient@example.com", "Person Data", "Please find attached the person data.", excelFile);

            return "Email sent successfully!";
        } catch (Exception e) {
            e.printStackTrace();
            return "Failed to send email!";
        }
    }
}
