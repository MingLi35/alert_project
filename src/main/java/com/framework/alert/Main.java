// Main.java
package com.framework.alert;

import com.framework.alert.model.FrameworkAgreement;
import com.framework.alert.model.MailContent;
import com.framework.alert.service.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.text.SimpleDateFormat;
import java.util.*;

public class Main {
    private static final Logger logger = LoggerFactory.getLogger(Main.class);
    private static final SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMdd_HHmmss");

    public static void main(String[] args) {
        try {
            logger.info("开始执行人月框架协议提醒程序");

            // 1. 设置Excel文件路径
            String excelPath = "2026年有时限要求事项清单.xlsx";
            if (args.length > 0) {
                excelPath = args[0];
            }

            File excelFile = new File(excelPath);
            if (!excelFile.exists()) {
                logger.error("Excel文件不存在: " + excelPath);
                System.err.println("请将Excel文件放置在: " + new File(".").getAbsolutePath());
                return;
            }

            logger.info("读取Excel文件: " + excelFile.getAbsolutePath());

            // 2. 读取Excel
            ExcelReader excelReader = new ExcelReader();
            List<FrameworkAgreement> agreements = excelReader.readExcel(excelPath);

            if (agreements.isEmpty()) {
                logger.error("没有读取到任何数据，请检查Excel文件格式");
                return;
            }

            // 3. 分析提醒
            AlertAnalyzer analyzer = new AlertAnalyzer();
            MailContent mailContent = analyzer.analyzeAlerts(agreements);

            // 4. 生成HTML
            EmailGenerator emailGenerator = new EmailGenerator();
            String htmlContent = emailGenerator.generateHtmlContent(mailContent);

            // 5. 保存HTML文件
            String timestamp = sdf.format(new Date());
            String outputPath = "" + timestamp + ".html";

            Files.write(Paths.get(outputPath), htmlContent.getBytes("UTF-8"));
            logger.info("提醒邮件已生成: " + outputPath);

            // 6. 在控制台输出摘要
            printSummary(mailContent);

            logger.info("程序执行完成");

        } catch (Exception e) {
            logger.error("程序执行失败", e);
            e.printStackTrace();
        }
    }

    private static void printSummary(MailContent mailContent) {
        // 用于重复字符的方法（替代Java 11的String.repeat()）
        String line = repeatString("=", 50);
        String dash = repeatString("-", 50);

        System.out.println("\n" + line);
        System.out.println("          人月框架协议提醒汇总");
        System.out.println(line);

        System.out.println("需整理资料: " + formatList(mailContent.getNeedDataAlert()));
        System.out.println("需完成事财权: " + formatList(mailContent.getNeedApprovalAlert()));
        System.out.println("需完成合同采购: " + formatList(mailContent.getNeedPurchaseAlert()));
        System.out.println("需完成合同用印: " + formatList(mailContent.getNeedSealAlert()));

        Set<String> allNames = new HashSet<>();
        allNames.addAll(mailContent.getNeedDataAlert());
        allNames.addAll(mailContent.getNeedApprovalAlert());
        allNames.addAll(mailContent.getNeedPurchaseAlert());
        allNames.addAll(mailContent.getNeedSealAlert());

        System.out.println(dash);
        System.out.println("总提醒人数: " + allNames.size());
        System.out.println("表格记录数: " + mailContent.getTableData().size());
        System.out.println(line + "\n");
    }

    private static String formatList(List<String> list) {
        if (list == null || list.isEmpty()) {
            return "无";
        }
        StringBuilder sb = new StringBuilder();
        for (int i = 0; i < list.size(); i++) {
            sb.append("@").append(list.get(i));
            if (i < list.size() - 1) {
                sb.append(", ");
            }
        }
        return sb.toString();
    }

    // 替代Java 11的String.repeat()方法
    private static String repeatString(String str, int count) {
        StringBuilder sb = new StringBuilder();
        for (int i = 0; i < count; i++) {
            sb.append(str);
        }
        return sb.toString();
    }
}