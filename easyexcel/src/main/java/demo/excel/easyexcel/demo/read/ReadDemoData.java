package demo.excel.easyexcel.demo.read;

import com.alibaba.excel.annotation.ExcelProperty;
import lombok.Data;

import java.util.Date;

@Data
public class ReadDemoData {
    /**===============================无注解===============================*/
     private String string;
     private Date date;
     private Double doubleData;

    // 无注解时，字段顺序必须与excel列顺序一致，否则解析的格式会
//    private Date date;
//    private String string;
//    private Double doubleData;


    /**===============================有注解===============================*/
    // 有注解时，excel也要有表头列，字段顺序可以不一致
//    @ExcelProperty("日期标题")
//    private Date date;
//    @ExcelProperty("数字标题")
//    private Double doubleData;
//    @ExcelProperty("字符串标题")
//    private String string;

}
