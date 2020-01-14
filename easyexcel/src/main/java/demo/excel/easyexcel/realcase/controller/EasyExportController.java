package demo.excel.easyexcel.realcase.controller;

import demo.excel.easyexcel.realcase.beans.DailyBaseData;
import demo.excel.easyexcel.realcase.beans.DailyMileData;
import demo.excel.easyexcel.realcase.beans.ParkingInfo;
import demo.excel.easyexcel.realcase.param.EasyExcelProperty;
import demo.excel.easyexcel.realcase.service.IBusinessService;
import demo.excel.easyexcel.realcase.service.IEasyExcelExportService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.util.List;

/**
 * 真实导出案例
 */
@RestController(value = "easyExport")
public class EasyExportController {

    @Autowired
    private IEasyExcelExportService easyExcelExportService; // 导出service
    @Autowired
    private IBusinessService businessService; // 业务service

    /**
     * 导出日报里程信息
     * @param request
     * @param response
     * @param vid   车辆id
     * @param stopDay   完成日期
     */
    @RequestMapping(value = "exportDailyMile", method = RequestMethod.GET)
    public void exportDailyMile(HttpServletRequest request, HttpServletResponse response, String vid, String stopDay) {
        // 构造数据
        DailyBaseData baseData = businessService.dailyBaseData();
        List<DailyMileData> mileDataList = businessService.dailyListData();
        // 填充到excel并导出
        easyExcelExportService.fillAndDownloadDailyDatas(response, baseData, mileDataList);
    }

    /**
     * 导出停车信息
     * @param request
     * @param response
     * @param vids
     * @throws Exception
     */
    @RequestMapping(value = "exportParkingInfo", method = RequestMethod.GET)
    public void exportParkingInfo(HttpServletRequest request, HttpServletResponse response, String vids) throws Exception{
        List<ParkingInfo> list = businessService.getParkingInfos(vids);
        EasyExcelProperty param = new EasyExcelProperty();
        param.setFileName("停车信息");
        param.setClazz(ParkingInfo.class);
        easyExcelExportService.simpleListExport(response, list, param);
    }

}
