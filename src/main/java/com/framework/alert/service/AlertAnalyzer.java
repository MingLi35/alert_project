package com.framework.alert.service;

import com.framework.alert.model.FrameworkAgreement;
import com.framework.alert.model.MailContent;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.*;
import java.util.stream.Collectors;

public class AlertAnalyzer {
    private static final Logger logger = LoggerFactory.getLogger(AlertAnalyzer.class);
    
    public MailContent analyzeAlerts(List<FrameworkAgreement> allAgreements) {
        // 过滤出运营业务开发科的记录
        List<FrameworkAgreement> targetAgreements = allAgreements.stream()
                .filter(a -> "运营业务开发科".equals(a.getResponsibleDepartment()))
                .collect(Collectors.toList());
        
        logger.info("找到 {} 条运营业务开发科的记录", targetAgreements.size());
        
        // 计算紧急程度
        calculateUrgencyLevel(targetAgreements);
        
        // 生成提醒
        MailContent mailContent = new MailContent();
        mailContent.setSubject("人月框架协议进度提醒 - " + new Date());
        mailContent.setNeedDataAlert(generateNeedDataAlert(targetAgreements));
        mailContent.setNeedApprovalAlert(generateNeedApprovalAlert(targetAgreements));
        mailContent.setNeedPurchaseAlert(generateNeedPurchaseAlert(targetAgreements));
        mailContent.setNeedSealAlert(generateNeedSealAlert(targetAgreements));
        
        // 排序表格数据（按紧急程度，已完成放最后）
        mailContent.setTableData(sortAgreements(targetAgreements));
        
        return mailContent;
    }
    
    private void calculateUrgencyLevel(List<FrameworkAgreement> agreements) {
        Date now = new Date();
        
        for (FrameworkAgreement agreement : agreements) {
            Date referenceDate = getReferenceDate(agreement);
            if (referenceDate == null) {
                agreement.setAlertLevel(5); // 最低优先级
                continue;
            }
            
            long diffDays = (referenceDate.getTime() - now.getTime()) / (1000 * 60 * 60 * 24);
            
            if (diffDays < 7) { // 7天内
                agreement.setAlertLevel(1);
            } else if (diffDays < 14) { // 7-14天
                agreement.setAlertLevel(2);
            } else if (diffDays < 30) { // 14-30天
                agreement.setAlertLevel(3);
            } else if (diffDays < 90) { // 30-90天
                agreement.setAlertLevel(4);
            } else {
                agreement.setAlertLevel(5);
            }
        }
    }
    
    private Date getReferenceDate(FrameworkAgreement agreement) {
        Date referenceDate = agreement.getPreviousAgreementExpiry();
        if (referenceDate == null) {
            referenceDate = agreement.getPlannedApprovalDate();
        }
        return referenceDate;
    }
    
    private List<String> generateNeedDataAlert(List<FrameworkAgreement> agreements) {
        Set<String> names = new HashSet<>();
        Date now = new Date();
        Calendar calendar = Calendar.getInstance();
        
        for (FrameworkAgreement agreement : agreements) {
            Date referenceDate = getReferenceDate(agreement);
            if (referenceDate == null) continue;
            
            calendar.setTime(referenceDate);
            calendar.add(Calendar.MONTH, -3);
            Date threeMonthsBefore = calendar.getTime();
            
            // 前三个月内
            if (now.after(threeMonthsBefore) && now.before(referenceDate)) {
                String progress = agreement.getCurrentProgress();
                if (progress == null || progress.isEmpty() || "资料整理中".equals(progress)) {
                    names.add(agreement.getResponsiblePerson());
                }
            }
        }
        return new ArrayList<>(names);
    }
    
    private List<String> generateNeedApprovalAlert(List<FrameworkAgreement> agreements) {
        Set<String> names = new HashSet<>();
        Date now = new Date();
        Calendar calendar = Calendar.getInstance();
        
        for (FrameworkAgreement agreement : agreements) {
            Date referenceDate = getReferenceDate(agreement);
            if (referenceDate == null) continue;
            
            calendar.setTime(referenceDate);
            calendar.add(Calendar.MONTH, -2);
            Date twoMonthsBefore = calendar.getTime();
            
            // 前两个月内
            if (now.after(twoMonthsBefore) && now.before(referenceDate)) {
                String progress = agreement.getCurrentProgress();
                if (progress == null || progress.isEmpty() || 
                    "资料整理中".equals(progress) || 
                    "发起事财权阶段".equals(progress)) {
                    names.add(agreement.getResponsiblePerson());
                }
            }
        }
        return new ArrayList<>(names);
    }
    
    private List<String> generateNeedPurchaseAlert(List<FrameworkAgreement> agreements) {
        Set<String> names = new HashSet<>();
        Date now = new Date();
        Calendar calendar = Calendar.getInstance();
        
        for (FrameworkAgreement agreement : agreements) {
            Date referenceDate = getReferenceDate(agreement);
            if (referenceDate == null) continue;
            
            calendar.setTime(referenceDate);
            calendar.add(Calendar.MONTH, -1);
            Date oneMonthBefore = calendar.getTime();
            
            // 前一个月内
            if (now.after(oneMonthBefore) && now.before(referenceDate)) {
                String progress = agreement.getCurrentProgress();
                if (progress == null || progress.isEmpty() || 
                    "资料整理中".equals(progress) || 
                    "发起事财权阶段".equals(progress) ||
                    "合同采购阶段".equals(progress)) {
                    names.add(agreement.getResponsiblePerson());
                }
            }
        }
        return new ArrayList<>(names);
    }
    
    private List<String> generateNeedSealAlert(List<FrameworkAgreement> agreements) {
        Set<String> names = new HashSet<>();
        Date now = new Date();
        Calendar calendar = Calendar.getInstance();
        
        for (FrameworkAgreement agreement : agreements) {
            Date referenceDate = getReferenceDate(agreement);
            if (referenceDate == null) continue;
            
            calendar.setTime(referenceDate);
            calendar.add(Calendar.DAY_OF_MONTH, -14);
            Date twoWeeksBefore = calendar.getTime();
            
            // 前两周内
            if (now.after(twoWeeksBefore) && now.before(referenceDate)) {
                String progress = agreement.getCurrentProgress();
                if (progress == null || progress.isEmpty() || 
                    "资料整理中".equals(progress) || 
                    "发起事财权阶段".equals(progress) ||
                    "合同采购阶段".equals(progress) ||
                    "合同用印阶段".equals(progress)) {
                    names.add(agreement.getResponsiblePerson());
                }
            }
        }
        return new ArrayList<>(names);
    }
    
    private List<FrameworkAgreement> sortAgreements(List<FrameworkAgreement> agreements) {
        return agreements.stream()
                .sorted(Comparator
                        .comparing((FrameworkAgreement a) -> "已完成".equals(a.getCurrentProgress()) ? 1 : 0)
                        .thenComparing(FrameworkAgreement::getAlertLevel))
                .collect(Collectors.toList());
    }
}