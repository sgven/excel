package demo.excel.easyexcel.realcase.beans;

import lombok.Data;

/**
 * 日报里程统计数据
 */
@Data
public class DailyMileData {
    private String testSiteId; // 区域id
    private String testSiteName; // 区域名称
    private Double doneMile; // 完成公里
    private Double totalMile; // 累计公里
    private Double targetMile; // 大纲要求（目标公里）
    private Double planDoneMile; // 计划完成公里
    private Double planTotalMile; // 计划累计公里
    private Double planTargetMile; // 计划大纲要求（目标公里）
}
