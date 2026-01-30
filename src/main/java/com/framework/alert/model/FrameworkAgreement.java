package com.framework.alert.model;

import lombok.Data;
import java.util.Date;

@Data
public class FrameworkAgreement {
    private Integer id;
    private String systemName;
    private String businessDepartment;
    private String responsibleDepartment;  // 责任科室
    private String frameworkAgreementName;
    private String projectDuration;  // 项目起止时间
    private Integer durationMonths;  // 签订期限(月)
    private String approvalFormNo;   // 《事财权审批表》单号
    private Integer plannedWorkload; // 拟立项需求工作量(人月)
    private Double plannedAmount;    // 2026年拟立项金额
    private Boolean isSigned;        // 是否签订框架协议/续约
    private Boolean involvesCentralizedPurchase; // 是否涉及集中采购
    private Date previousAgreementExpiry; // 上期协议到期
    private Date plannedApprovalDate;    // 计划完成立项日期
    private Date plannedPurchaseDate;    // 计划完成采购日期
    private Date plannedContractDate;    // 计划合同签订日期
    private String responsiblePerson;    // 责任经办
    private String currentProgress;      // 当前进度
    private Boolean hasOverdueRisk;      // 是否存在逾期风险/已逾期
    private String riskDescription;      // 风险情况说明/逾期情况
    private String countermeasures;      // 风险应对措施/需中心协助事项
    private String remarks;
    
    // 辅助字段
    private Date referenceDate;          // 参考日期（根据规则确定）
    private Integer alertLevel;          // 紧急程度
}