package com.framework.alert.model;

import lombok.Data;
import java.util.List;

@Data
public class MailContent {
    private String subject;
    private List<String> needDataAlert;          // 需整理资料提醒
    private List<String> needApprovalAlert;      // 需完成事财权提醒
    private List<String> needPurchaseAlert;      // 需完成合同采购提醒
    private List<String> needSealAlert;          // 需完成合同用印提醒
    private List<FrameworkAgreement> tableData;  // 表格数据
}