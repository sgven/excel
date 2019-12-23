package demo.excel.easyexcel.realcase.beans;

import com.alibaba.excel.annotation.ExcelIgnore;
import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.write.style.ColumnWidth;

/**
 * 停车信息
 */
public class ParkingInfo {

    @ExcelIgnore
    private String orderNo; // 序号
    @ExcelProperty(value = "样车型号", index = 0)
    private String vehicleType; // 样车型号
    @ColumnWidth(10)
    @ExcelProperty(value = "VIN", index = 1)
    private String vinNo; // VIN号
    @ExcelProperty(value = "停放位置", index = 2)
    private String parkingPlace; // 停放位置
    @ExcelProperty(value = "停放时长", index = 3)
    private String parkingTime; // 停放时长
    @ExcelProperty(value = "试验员", index = 4)
    private String testOperator; // 试验员
    @ExcelProperty(value = "试验经理", index = 5)
    private String testManager; // 试验经理

    public String getOrderNo() {
        return orderNo;
    }

    public void setOrderNo(String orderNo) {
        this.orderNo = orderNo;
    }

    public String getVehicleType() {
        return vehicleType;
    }

    public void setVehicleType(String vehicleType) {
        this.vehicleType = vehicleType;
    }

    public String getVinNo() {
        return vinNo;
    }

    public void setVinNo(String vinNo) {
        this.vinNo = vinNo;
    }

    public String getParkingPlace() {
        return parkingPlace;
    }

    public void setParkingPlace(String parkingPlace) {
        this.parkingPlace = parkingPlace;
    }

    public String getParkingTime() {
        return parkingTime;
    }

    public void setParkingTime(String parkingTime) {
        this.parkingTime = parkingTime;
    }

    public String getTestOperator() {
        return testOperator;
    }

    public void setTestOperator(String testOperator) {
        this.testOperator = testOperator;
    }

    public String getTestManager() {
        return testManager;
    }

    public void setTestManager(String testManager) {
        this.testManager = testManager;
    }
}
