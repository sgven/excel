package demo.excel.easyexcel.realcase.service.impl;

import demo.excel.easyexcel.realcase.beans.DailyBaseData;
import demo.excel.easyexcel.realcase.beans.DailyMileData;
import demo.excel.easyexcel.realcase.beans.ParkingInfo;
import demo.excel.easyexcel.realcase.service.IBusinessService;
import org.springframework.stereotype.Service;

import java.util.ArrayList;
import java.util.List;

@Service
public class BusinessServiceImpl implements IBusinessService {
    @Override
    public DailyBaseData dailyBaseData() {
        return demoBaseData();
    }

    @Override
    public List<DailyMileData> dailyListData() {
        return demoListData();
    }

    @Override
    public List<ParkingInfo> getParkingInfos(String vids) {
        return demoSimpleList();
    }

    private List<ParkingInfo> demoSimpleList() {
        List<ParkingInfo> list = new ArrayList<ParkingInfo>();
        for(int i = 0; i < 5; i++) {
            list.add(demoParkingInfo());
        }
        return list;
    }

    private ParkingInfo demoParkingInfo() {
        ParkingInfo data = new ParkingInfo();
        data.setParkingTime("5");
        data.setParkingPlace("万达地下停车场");
        data.setVinNo("0001");
        return data;
    }

    private DailyBaseData demoBaseData() {
        DailyBaseData myFillData = new DailyBaseData();
        myFillData.setVinNo("H02C-4#(000037)");
        myFillData.setDoneDate("2019.12.19");
        myFillData.setPlanDate("2019.12.20");
        myFillData.setTestNodeTime("2019.12.31");
        myFillData.setPreDoneTime("2019.12.31");
        myFillData.setStatusName("正常");
        return myFillData;
    }

    private List<DailyMileData> demoListData() {
        List<DailyMileData> list = new ArrayList<DailyMileData>();
        list.add(testData());
        list.add(testData());
        list.add(testData());
        list.add(testData());
        return list;
    }

    private DailyMileData testData() {
        DailyMileData testData = new DailyMileData();
        testData.setDoneMile(500.0);
        testData.setTotalMile(6315.0);
        testData.setTargetMile(6300.0);
        testData.setPlanDoneMile(0.0);
        testData.setPlanTotalMile(6300.0);
        testData.setPlanTargetMile(6300.0);
        return testData;
    }
}
