package demo.excel.easyexcel.realcase.service.impl;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.write.builder.ExcelWriterBuilder;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.alibaba.excel.write.metadata.fill.FillConfig;
import com.alibaba.excel.write.style.column.LongestMatchColumnWidthStyleStrategy;
import demo.excel.easyexcel.realcase.beans.DailyBaseData;
import demo.excel.easyexcel.realcase.beans.DailyMileData;
import demo.excel.easyexcel.realcase.handler.DefaultHeadHandler;
import demo.excel.easyexcel.realcase.param.EasyExcelProperty;
import demo.excel.easyexcel.realcase.service.IEasyExcelExportService;
import org.springframework.stereotype.Service;
import org.springframework.util.Assert;

import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.net.URLEncoder;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@Service
public class EasyExcelExportServiceImpl implements IEasyExcelExportService {

    @Override
    public void fillAndDownloadDailyDatas(HttpServletResponse response, DailyBaseData baseInfo, List<DailyMileData> mileDatas) {
        try {
            // 模板文件
            String templateFileName =
                    getClass().getResource("/").getPath() + "templates" + File.separator + "daily_mile_template.xlsx";

            response.setContentType("application/vnd.ms-excel");
            response.setCharacterEncoding("utf-8");
            // URLEncoder.encode可以防止中文乱码
            String fileName = URLEncoder.encode("日报里程信息", "UTF-8");
            response.setHeader("Content-disposition", "attachment;filename=" + fileName + ".xlsx");

            ExcelWriterBuilder excelWriterBuilder = EasyExcel.write(response.getOutputStream())
                    .autoCloseStream(Boolean.FALSE) // 这里需要设置不关闭流
                    .withTemplate(templateFileName); // 使用模板
            ExcelWriter excelWriter = excelWriterBuilder.build();
            WriteSheet writeSheet = EasyExcel.writerSheet().sheetName("日报里程信息").build();
            excelWriter.fill(baseInfo, writeSheet);
            // 填充list数据,false会直接使用下一行，如果没有则创建；如果true，则下面有没有空行 都会创建一行
            FillConfig fillConfig = FillConfig.builder().forceNewRow(Boolean.FALSE).build();
            excelWriter.fill(mileDatas, fillConfig, writeSheet);
            // 千万别忘记finish 会帮忙关闭流
            excelWriter.finish();
        } catch (Exception e) {
            // 重置response
            response.reset();
            response.setContentType("application/json");
            response.setCharacterEncoding("utf-8");
            Map<String, String> map = new HashMap<String, String>();
            map.put("status", "failure");
            map.put("message", "下载文件失败" + e.getMessage());
//            response.getWriter().println(JSON.toJSONString(map));
        }
    }

    @Override
    public void simpleListExport(HttpServletResponse response, List list, EasyExcelProperty param) {
        try {
            String filename = param.getFileName();
            String titleName = param.getTitleName();
            String sheetName = param.getSheetName();
            Class clazz = param.getClazz();
            Assert.notNull(filename, "文件名不能为空");
            Assert.notNull(clazz, "导出实体类不能为空");

            response.setContentType("application/vnd.ms-excel");
            response.setCharacterEncoding("utf-8");
            // URLEncoder.encode可以防止中文乱码
            String fileName = URLEncoder.encode(filename, "UTF-8");
            response.setHeader("Content-disposition", "attachment;filename=" + fileName + ".xlsx");

            ExcelWriterBuilder excelWriterBuilder = EasyExcel.write(response.getOutputStream(), clazz)
                    .registerWriteHandler(new LongestMatchColumnWidthStyleStrategy()) // 自动列宽（不太精确）
                    .autoCloseStream(Boolean.FALSE); // 这里需要设置不关闭流

            new DefaultHeadHandler().handle(excelWriterBuilder, param);

            ExcelWriter excelWriter = excelWriterBuilder.build();
            WriteSheet writeSheet = EasyExcel.writerSheet().build();
            excelWriter.write(list, writeSheet);
            // 千万别忘记finish 会帮忙关闭流
            excelWriter.finish();
        } catch (Exception e) {
            // 重置response
            response.reset();
            response.setContentType("application/json");
            response.setCharacterEncoding("utf-8");
            Map<String, String> map = new HashMap<String, String>();
            map.put("status", "failure");
            map.put("message", "下载文件失败" + e.getMessage());
//            response.getWriter().println(JSON.toJSONString(map));
        }
    }

    /**
     * web上传、下载 和 本地文件的读写 "区别"仅仅是:
     * EasyExcel.write()时, 指定的输入输出是 response的输出流，还是filename。
     * 当然还有一点细微的差别，如下代码所示。
     *
     * 如果是web上传、下载（通过浏览器下载），则 EasyExcel.write()时指定一个流（response.输出流）
     * 如果是导出到本地（指定路径的）文件，则 EasyExcel.write()时指定一个文件名
     * @param response
     * @param baseInfo
     * @param mileDatas
     */
    private void fillDailyDatas(HttpServletResponse response, DailyBaseData baseInfo, List<DailyMileData> mileDatas) {
        // 模板文件
        String templateFileName =
                getClass().getResource("/").getPath() + "templates" + File.separator + "daily_mile_template.xlsx";

        String fileName = getClass().getResource("/").getPath() + "日报里程信息" + System.currentTimeMillis() + ".xlsx";

        // web上传、下载
        /*
        response.setHeader("Content-disposition", "attachment;filename=" + fileName + ".xlsx");
        ExcelWriterBuilder excelWriterBuilder = EasyExcel.write(response.getOutputStream())
                .autoCloseStream(Boolean.FALSE) // 这里需要设置不关闭流
                .withTemplate(templateFileName);
        ExcelWriter excelWriter = excelWriterBuilder.build();
        ExcelWriter excelWriter = EasyExcel.write(response.getOutputStream())..autoCloseStream(Boolean.FALSE).withTemplate(templateFileName).build();*/
        // 本地文件的读写
       /* ExcelWriterBuilder excelWriterBuilder = EasyExcel.write(fileName)
                .withTemplate(templateFileName);
        ExcelWriter excelWriter = excelWriterBuilder.build();*/
        ExcelWriter excelWriter = EasyExcel.write(fileName).withTemplate(templateFileName).build();

        WriteSheet writeSheet = EasyExcel.writerSheet().build();
        excelWriter.fill(baseInfo, writeSheet);
        FillConfig fillConfig = FillConfig.builder().forceNewRow(Boolean.FALSE).build();
        excelWriter.fill(mileDatas, fillConfig, writeSheet);
        // 千万别忘记finish 会帮忙关闭流
        excelWriter.finish();
    }

}
