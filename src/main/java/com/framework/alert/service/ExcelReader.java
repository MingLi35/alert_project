package com.framework.alert.service;

import com.framework.alert.model.FrameworkAgreement;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.FileInputStream;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

public class ExcelReader {
    private static final Logger logger = LoggerFactory.getLogger(ExcelReader.class);
    private static final SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
    private static final SimpleDateFormat sdf2 = new SimpleDateFormat("yyyy/MM/dd");
    
    public List<FrameworkAgreement> readExcel(String filePath) throws Exception {
        List<FrameworkAgreement> agreements = new ArrayList<>();
        
        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(fis)) {
            
            Sheet sheet = workbook.getSheetAt(1); // 第二个sheet
            
            for (int i = 2; i <= sheet.getLastRowNum(); i++) { // 从第3行开始
                Row row = sheet.getRow(i);
                if (row == null) continue;
                
                FrameworkAgreement agreement = parseRow(row);
                if (agreement != null) {
                    agreements.add(agreement);
                }
            }
        }
        
        logger.info("成功读取 {} 条记录", agreements.size());
        return agreements;
    }
    
    private FrameworkAgreement parseRow(Row row) {
        try {
            FrameworkAgreement agreement = new FrameworkAgreement();
            
            // 序号 (A列)
            agreement.setId(getIntValue(row.getCell(0)));
            
            // 系统名称 (B列)
            agreement.setSystemName(getStringValue(row.getCell(1)));
            
            // 业务归口管理部门 (C列)
            agreement.setBusinessDepartment(getStringValue(row.getCell(2)));
            
            // 责任科室 (D列)
            agreement.setResponsibleDepartment(getStringValue(row.getCell(3)));
            
            // 责任经办 (P列 - 注意Excel索引从0开始)
            agreement.setResponsiblePerson(getStringValue(row.getCell(15)));
            
            // 当前进度 (Q列)
            agreement.setCurrentProgress(getStringValue(row.getCell(16)));
            
            // 上期协议到期 (L列)
            agreement.setPreviousAgreementExpiry(parseDateCell(row.getCell(11)));
            
            // 计划完成立项日期 (M列)
            agreement.setPlannedApprovalDate(parseDateCell(row.getCell(12)));
            
            // 计划完成采购日期 (N列)
            agreement.setPlannedPurchaseDate(parseDateCell(row.getCell(13)));
            
            // 计划合同签订日期 (O列)
            agreement.setPlannedContractDate(parseDateCell(row.getCell(14)));
            
            return agreement;
            
        } catch (Exception e) {
            logger.error("解析Excel行数据失败，行号: " + row.getRowNum(), e);
            return null;
        }
    }
    
    private String getStringValue(Cell cell) {
        if (cell == null) return "";
        
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue().trim();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return sdf.format(cell.getDateCellValue());
                }
                // 如果是整数，去掉小数部分
                double num = cell.getNumericCellValue();
                if (num == (int) num) {
                    return String.valueOf((int) num);
                }
                return String.valueOf(num);
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                try {
                    return cell.getStringCellValue();
                } catch (Exception e) {
                    try {
                        return String.valueOf(cell.getNumericCellValue());
                    } catch (Exception e2) {
                        return "";
                    }
                }
            default:
                return "";
        }
    }
    
    private Integer getIntValue(Cell cell) {
        if (cell == null) return 0;
        
        if (cell.getCellType() == CellType.NUMERIC) {
            return (int) cell.getNumericCellValue();
        } else if (cell.getCellType() == CellType.STRING) {
            try {
                return Integer.parseInt(cell.getStringCellValue().trim());
            } catch (NumberFormatException e) {
                return 0;
            }
        }
        return 0;
    }
    
    private Date parseDateCell(Cell cell) {
        if (cell == null) return null;
        
        try {
            if (cell.getCellType() == CellType.NUMERIC && DateUtil.isCellDateFormatted(cell)) {
                return cell.getDateCellValue();
            } else if (cell.getCellType() == CellType.STRING) {
                String dateStr = cell.getStringCellValue().trim();
                if (dateStr.isEmpty() || "上期未签订".equals(dateStr)) {
                    return null;
                }
                
                // 尝试不同的日期格式
                try {
                    return sdf.parse(dateStr);
                } catch (ParseException e1) {
                    try {
                        return sdf2.parse(dateStr);
                    } catch (ParseException e2) {
                        // 尝试其他格式
                        String[] formats = {
                            "yyyy-MM-dd HH:mm:ss",
                            "yyyy/MM/dd HH:mm:ss",
                            "yyyy年MM月dd日"
                        };
                        for (String format : formats) {
                            try {
                                return new SimpleDateFormat(format).parse(dateStr);
                            } catch (ParseException e) {
                                // 继续尝试
                            }
                        }
                        logger.warn("无法解析日期: " + dateStr);
                        return null;
                    }
                }
            }
        } catch (Exception e) {
            logger.warn("解析日期失败: " + cell, e);
        }
        return null;
    }
}