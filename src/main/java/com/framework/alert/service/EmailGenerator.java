package com.framework.alert.service;

import com.framework.alert.model.FrameworkAgreement;
import com.framework.alert.model.MailContent;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;

public class EmailGenerator {
    private static final Logger logger = LoggerFactory.getLogger(EmailGenerator.class);
    private static final SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
    
    public String generateHtmlContent(MailContent mailContent) {
        StringBuilder html = new StringBuilder();

        html.append("<html>");
        html.append("<head>");
        html.append("<style>");
        html.append("body { font-family: 'Microsoft YaHei', Arial, sans-serif; margin: 20px; line-height: 1.6; }");
        html.append(".alert-section { margin: 20px 0; padding: 15px; background-color: #f8f9fa; border-left: 4px solid #007bff; border-radius: 4px; }");
        html.append(".alert-title { font-weight: bold; color: #333; margin-bottom: 10px; font-size: 16px; }");
        html.append(".alert-content { color: #666; }");
        html.append(".person-name { color: #e74c3c; font-weight: bold; }");
        html.append("table { border-collapse: collapse; width: 100%; margin-top: 30px; box-shadow: 0 2px 5px rgba(0,0,0,0.1); }");
        html.append("th, td { border: 1px solid #ddd; padding: 12px; text-align: left; }");
        html.append("th { background-color: #2c3e50; color: white; font-weight: bold; }");
        html.append("tr:nth-child(even) { background-color: #f9f9f9; }");
        html.append("tr:hover { background-color: #f5f5f5; }");
        html.append(".urgent { background-color: #ffe6e6 !important; }");
        html.append(".completed { background-color: #e6ffe6 !important; }");
        html.append(".level-1 { color: #e74c3c; font-weight: bold; }");
        html.append(".level-2 { color: #e67e22; }");
        html.append(".level-3 { color: #f1c40f; }");
        html.append(".level-4 { color: #3498db; }");
        html.append(".level-5 { color: #95a5a6; }");
        html.append(".header { background-color: #34495e; color: white; padding: 20px; border-radius: 5px; margin-bottom: 20px; }");
        html.append(".footer { margin-top: 30px; padding-top: 20px; border-top: 1px solid #eee; color: #7f8c8d; font-size: 12px; }");
        html.append("</style>");
        html.append("</head>");
        html.append("<body>");

        // 头部
        html.append("<div class=\"header\">");
        html.append("<h1>人月框架协议进度提醒</h1>");
        html.append("<p>生成时间: ").append(new Date()).append("</p>");
        html.append("</div>");

        // 提醒部分
        List<String> needDataAlert = mailContent.getNeedDataAlert();
        List<String> needApprovalAlert = mailContent.getNeedApprovalAlert();
        List<String> needPurchaseAlert = mailContent.getNeedPurchaseAlert();
        List<String> needSealAlert = mailContent.getNeedSealAlert();

        if (!needDataAlert.isEmpty() || !needApprovalAlert.isEmpty() ||
                !needPurchaseAlert.isEmpty() || !needSealAlert.isEmpty()) {
            html.append("<h2>📢 待办事项提醒</h2>");

            addAlertSection(html, "📋 以下同事请及时整理人月框架协议资料：", needDataAlert);
            addAlertSection(html, "📝 以下同事请及时完成事财权审批：", needApprovalAlert);
            addAlertSection(html, "🛒 以下同事请及时完成合同采购：", needPurchaseAlert);
            addAlertSection(html, "🖋️ 以下同事请及时完成合同用印：", needSealAlert);
        } else {
            html.append("<div class=\"alert-section\">");
            html.append("<div class=\"alert-title\">✅ 所有事项正常</div>");
            html.append("<div class=\"alert-content\">目前没有需要紧急处理的事项。</div>");
            html.append("</div>");
        }

        // 表格部分
        List<FrameworkAgreement> tableData = mailContent.getTableData();
        if (tableData != null && !tableData.isEmpty()) {
            html.append("<h2>📊 运营业务开发科项目清单（共").append(tableData.size()).append("项）</h2>");
            html.append("<table>");
            html.append("<tr>");
            html.append("<th width=\"5%\">序号</th>");
            html.append("<th width=\"25%\">系统名称</th>");
            html.append("<th width=\"10%\">责任经办</th>");
            html.append("<th width=\"15%\">当前进度</th>");
            html.append("<th width=\"15%\">上期协议到期</th>");
            html.append("<th width=\"15%\">计划立项日期</th>");
            html.append("<th width=\"15%\">紧急程度</th>");
            html.append("</tr>");

            for (FrameworkAgreement agreement : tableData) {
                String rowClass = "";
                String currentProgress = agreement.getCurrentProgress();

                if (currentProgress != null && currentProgress.contains("已完成")) {
                    rowClass = "completed";
                } else if (agreement.getAlertLevel() != null && agreement.getAlertLevel() <= 2) {
                    rowClass = "urgent";
                }

                html.append("<tr class=\"").append(rowClass).append("\">");
                html.append("<td>").append(agreement.getId() != null ? agreement.getId() : "").append("</td>");
                html.append("<td>").append(agreement.getSystemName() != null ? agreement.getSystemName() : "").append("</td>");
                html.append("<td>").append(agreement.getResponsiblePerson() != null ?
                        agreement.getResponsiblePerson() : "").append("</td>");
                html.append("<td>").append(currentProgress != null ? currentProgress : "").append("</td>");
                html.append("<td>").append(formatDate(agreement.getPreviousAgreementExpiry())).append("</td>");
                html.append("<td>").append(formatDate(agreement.getPlannedApprovalDate())).append("</td>");
                html.append("<td class=\"level-").append(agreement.getAlertLevel() != null ?
                        agreement.getAlertLevel() : 5).append("\">");
                html.append(getUrgencyText(agreement.getAlertLevel())).append("</td>");
                html.append("</tr>");
            }

            html.append("</table>");

            // 统计信息
            long completedCount = tableData.stream()
                    .filter(a -> a.getCurrentProgress() != null && a.getCurrentProgress().contains("已完成"))
                    .count();
            long urgentCount = tableData.stream()
                    .filter(a -> a.getAlertLevel() != null && a.getAlertLevel() <= 2)
                    .count();

            html.append("<div class=\"footer\">");
            html.append("<p>📈 统计信息：总计 ").append(tableData.size()).append(" 项");
            html.append(" | 已完成 ").append(completedCount).append(" 项");
            html.append(" | 紧急待办 ").append(urgentCount).append(" 项</p>");
            html.append("<p>📋 紧急程度说明：非常紧急(7天内) | 紧急(7-14天) | 中等(14-30天) | 一般(30-90天) | 较低(90天以上)</p>");
            html.append("<p>✅ 绿色行：已完成项目 | 🔴 红色背景：紧急项目</p>");
            html.append("</div>");
        } else {
            html.append("<div class=\"alert-section\">");
            html.append("<div class=\"alert-title\">⚠️ 未找到相关记录</div>");
            html.append("<div class=\"alert-content\">未找到'运营业务开发科'的相关项目记录。</div>");
            html.append("</div>");
        }

        html.append("</body>");
        html.append("</html>");

        return html.toString();
    }
    
    private void addAlertSection(StringBuilder html, String title, List<String> names) {
        if (!names.isEmpty()) {
            html.append("<div class=\"alert-section\">");
            html.append("<div class=\"alert-title\">").append(title).append("</div>");
            html.append("<div class=\"alert-content\">");
            for (int i = 0; i < names.size(); i++) {
                html.append("<span class=\"person-name\">@").append(names.get(i)).append("</span>");
                if (i < names.size() - 1) {
                    html.append(", ");
                }
            }
            html.append("</div>");
            html.append("</div>");
        }
    }
    
    private String formatDate(Date date) {
        return date != null ? sdf.format(date) : "-";
    }
    
    private String getUrgencyText(Integer level) {
        if (level == null) return "较低";
        switch (level) {
            case 1: return "非常紧急";
            case 2: return "紧急";
            case 3: return "中等";
            case 4: return "一般";
            case 5: return "较低";
            default: return "较低";
        }
    }
}