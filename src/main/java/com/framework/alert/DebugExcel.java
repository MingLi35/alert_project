// DebugExcel.java
package com.framework.alert;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.text.SimpleDateFormat;

public class DebugExcel {
    public static void main(String[] args) {
        try {
            String filePath = "2026年有时限要求事项清单.xlsx";
            
            System.out.println("=== Excel文件结构分析 ===");
            System.out.println("文件路径: " + filePath);
            
            try (FileInputStream fis = new FileInputStream(filePath);
                 Workbook workbook = new XSSFWorkbook(fis)) {
                
                // 检查所有sheet
                System.out.println("\n=== Sheet列表 ===");
                for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                    Sheet sheet = workbook.getSheetAt(i);
                    System.out.println("Sheet " + i + ": " + sheet.getSheetName() + 
                                     " (行数: " + sheet.getLastRowNum() + ")");
                }
                
                // 分析第二个sheet
                System.out.println("\n=== 分析第二个Sheet: " + workbook.getSheetName(1) + " ===");
                Sheet sheet = workbook.getSheetAt(1);
                
                // 打印前5行
                System.out.println("\n=== 前5行内容 ===");
                for (int rowNum = 0; rowNum <= 4 && rowNum <= sheet.getLastRowNum(); rowNum++) {
                    Row row = sheet.getRow(rowNum);
                    if (row != null) {
                        System.out.println("\n--- 第" + (rowNum + 1) + "行 ---");
                        for (int colNum = 0; colNum < row.getLastCellNum(); colNum++) {
                            Cell cell = row.getCell(colNum);
                            String value = getCellValue(cell);
                            if (!value.isEmpty()) {
                                System.out.println("  列" + colNum + " (" + getColumnLetter(colNum) + "): " + value);
                            }
                        }
                    }
                }
                
                // 特别关注重要列
                System.out.println("\n=== 重要列定位 ===");
                Row headerRow = sheet.getRow(2); // 假设表头在第3行
                if (headerRow != null) {
                    // 查找"责任科室"列
                    findColumn(headerRow, "责任科室");
                    // 查找"责任经办"列
                    findColumn(headerRow, "责任经办");
                    // 查找"当前进度"列
                    findColumn(headerRow, "当前进度");
                    // 查找"上期协议到期"列
                    findColumn(headerRow, "上期协议到期");
                }
                
                // 查找"运营业务开发科"的记录
                System.out.println("\n=== 查找'运营业务开发科'记录 ===");
                int foundCount = 0;
                for (int rowNum = 3; rowNum <= sheet.getLastRowNum(); rowNum++) {
                    Row row = sheet.getRow(rowNum);
                    if (row != null) {
                        Cell deptCell = row.getCell(3); // 假设责任科室在第4列
                        if (deptCell != null && "运营业务开发科".equals(getCellValue(deptCell))) {
                            foundCount++;
                            System.out.println("\n找到记录 #" + foundCount + " (行" + (rowNum + 1) + "):");
                            System.out.println("  序号: " + getCellValue(row.getCell(0)));
                            System.out.println("  系统名称: " + getCellValue(row.getCell(1)));
                            System.out.println("  责任科室: " + getCellValue(deptCell));
                            
                            // 尝试不同的列索引查找责任经办
                            for (int col = 10; col <= 20; col++) {
                                String value = getCellValue(row.getCell(col));
                                if (value != null && !value.isEmpty() && 
                                    (value.contains("张") || value.contains("李") || 
                                     value.contains("王") || value.contains("赵") ||
                                     value.contains("刘") || value.contains("陈"))) {
                                    System.out.println("  责任经办(可能列" + col + "): " + value);
                                }
                            }
                        }
                    }
                }
                
                if (foundCount == 0) {
                    System.out.println("未找到'运营业务开发科'的记录！");
                    System.out.println("\n=== 所有责任科室值 ===");
                    for (int rowNum = 3; rowNum <= Math.min(20, sheet.getLastRowNum()); rowNum++) {
                        Row row = sheet.getRow(rowNum);
                        if (row != null) {
                            String dept = getCellValue(row.getCell(3));
                            if (!dept.isEmpty()) {
                                System.out.println("行" + (rowNum + 1) + ": " + dept);
                            }
                        }
                    }
                }
                
            }
            
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
    
    private static void findColumn(Row headerRow, String columnName) {
        for (int i = 0; i < headerRow.getLastCellNum(); i++) {
            Cell cell = headerRow.getCell(i);
            String value = getCellValue(cell);
            if (value.contains(columnName)) {
                System.out.println("找到'" + columnName + "': 列" + i + " (" + getColumnLetter(i) + ")");
                return;
            }
        }
        System.out.println("未找到'" + columnName + "'列");
    }
    
    private static String getCellValue(Cell cell) {
        if (cell == null) return "";
        
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue().trim();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return new SimpleDateFormat("yyyy-MM-dd").format(cell.getDateCellValue());
                }
                return String.valueOf(cell.getNumericCellValue());
            default:
                return "";
        }
    }
    
    private static String getColumnLetter(int index) {
        StringBuilder result = new StringBuilder();
        while (index >= 0) {
            result.insert(0, (char) ('A' + index % 26));
            index = index / 26 - 1;
        }
        return result.toString();
    }
}