package demo.excel.easyexcel.realcase.beans;

import lombok.Data;

/**
 * 日报里程基本信息
 */
@Data
public class DailyBaseData {
    private String vinNo; // 车型编号
    private String doneDate; //完成日期
    private String planDate; //计划日期
    private String testNodeTime; // 试验节点时间
    private String preDoneTime; // 预计完成时间
    private String status; // 试验进度
    private String statusName; // 试验进度名称
}
