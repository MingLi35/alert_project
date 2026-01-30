// ExcelReader.java - 修正列索引版本
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

            // 序号 (A列, 索引0)
            agreement.setId(getIntValue(row.getCell(0)));

            // 系统名称 (B列, 索引1)
            agreement.setSystemName(getStringValue(row.getCell(1)));

            // 业务归口管理部门 (C列, 索引2)
            agreement.setBusinessDepartment(getStringValue(row.getCell(2)));

            // 责任科室 (D列, 索引3) - 关键过滤列
            String responsibleDept = getStringValue(row.getCell(3));
            agreement.setResponsibleDepartment(responsibleDept);

            // 责任经办 (Q列, 索引16) ⭐ 修正：从15改为16
            agreement.setResponsiblePerson(getStringValue(row.getCell(16)));

            // 当前进度 (R列, 索引17) ⭐ 修正：从16改为17
            agreement.setCurrentProgress(getStringValue(row.getCell(17)));

            // 上期协议到期 (M列, 索引12) ⭐ 修正：从11改为12
            agreement.setPreviousAgreementExpiry(parseDateCell(row.getCell(12)));

            // 计划完成立项日期 (N列, 索引13)
            agreement.setPlannedApprovalDate(parseDateCell(row.getCell(13)));

            // 计划完成采购日期 (O列, 索引14)
            agreement.setPlannedPurchaseDate(parseDateCell(row.getCell(14)));

            // 计划合同签订日期 (P列, 索引15)
            agreement.setPlannedContractDate(parseDateCell(row.getCell(15)));

            // 调试输出运营业务开发科的记录
            if ("运营业务开发科".equals(responsibleDept)) {
                logger.debug("运营业务开发科记录: ID={}, 系统={}, 经办={}, 进度={}",
                        agreement.getId(), agreement.getSystemName(),
                        agreement.getResponsiblePerson(), agreement.getCurrentProgress());
            }

            return agreement;

        } catch (Exception e) {
            logger.error("解析Excel行数据失败，行号: {}", row.getRowNum() + 1, e);
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
            // 1. 处理数字格式的日期
            if (cell.getCellType() == CellType.NUMERIC && DateUtil.isCellDateFormatted(cell)) {
                return cell.getDateCellValue();
            }
            // 2. 处理字符串格式的日期
            else if (cell.getCellType() == CellType.STRING) {
                String dateStr = cell.getStringCellValue().trim();
                if (dateStr.isEmpty() || "上期未签订".equals(dateStr) || "N/A".equals(dateStr)) {
                    return null;
                }

                // 处理日期字符串
                return parseDateString(dateStr);
            }
            // 3. 处理可能是公式的单元格
            else if (cell.getCellType() == CellType.FORMULA) {
                try {
                    if (DateUtil.isCellDateFormatted(cell)) {
                        return cell.getDateCellValue();
                    }
                } catch (Exception e) {
                    // 忽略，尝试其他方式
                }
            }
        } catch (Exception e) {
            logger.warn("解析日期单元格失败: {}", cell, e);
        }
        return null;
    }

    private Date parseDateString(String dateStr) {
        if (dateStr == null || dateStr.trim().isEmpty()) {
            return null;
        }

        // 支持的日期格式
        String[] dateFormats = {
                "yyyy-MM-dd HH:mm:ss",
                "yyyy-MM-dd",
                "yyyy/MM/dd HH:mm:ss",
                "yyyy/MM/dd",
                "yyyy年MM月dd日",
                "yyyy.MM.dd"
        };

        for (String format : dateFormats) {
            try {
                SimpleDateFormat sdf = new SimpleDateFormat(format);
                sdf.setLenient(false); // 严格模式
                return sdf.parse(dateStr);
            } catch (ParseException e) {
                // 继续尝试下一个格式
            }
        }

        logger.warn("无法解析日期字符串: {}", dateStr);
        return null;
    }
}