package demo.excel.easyexcel.realcase.service;

import demo.excel.easyexcel.realcase.beans.DailyBaseData;
import demo.excel.easyexcel.realcase.beans.DailyMileData;
import demo.excel.easyexcel.realcase.beans.ParkingInfo;

import java.util.List;

public interface IBusinessService {
    /**
     * 日报里程基本数据
     * @return
     */
    DailyBaseData dailyBaseData();

    /**
     * 日报里程list数据（里程统计信息）
     * @return
     */
    List<DailyMileData> dailyListData();

    /**
     * 停车信息
     * @param vids
     * @return
     */
    List<ParkingInfo> getParkingInfos(String vids);
}
