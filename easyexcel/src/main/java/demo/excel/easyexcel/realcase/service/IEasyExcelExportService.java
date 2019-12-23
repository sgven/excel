package demo.excel.easyexcel.realcase.service;

import demo.excel.easyexcel.realcase.beans.DailyBaseData;
import demo.excel.easyexcel.realcase.beans.DailyMileData;

import javax.servlet.http.HttpServletResponse;
import java.util.List;

public interface IEasyExcelExportService {

    /**
     * 填充日报里程信息并 “导出”——（通过浏览器下载）
     * @param response
     * @param baseInfo
     * @param mileDatas
     */
    void fillAndDownloadDailyDatas(HttpServletResponse response, DailyBaseData baseInfo, List<DailyMileData> mileDatas);

    /**
     * 简单list “导出”——（通过浏览器下载）
     * @param response
     * @param list
     * @param filename
     */
    void simpleListExport(HttpServletResponse response, String filename, List list, Class clazz);
}
