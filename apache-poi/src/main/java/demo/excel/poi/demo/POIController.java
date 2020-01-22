package demo.excel.poi.demo;

import demo.excel.poi.util.ExcelUtil;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.FileInputStream;
import java.net.URLEncoder;
import java.util.ArrayList;
import java.util.List;

@RestController
@RequestMapping("poi")
public class POIController {

    private static final Logger logger = LoggerFactory.getLogger(ExcelUtil.class);

    @RequestMapping("merge")
    public void merge(HttpServletRequest request, HttpServletResponse response) throws Exception {
        response.setContentType("application/vnd.ms-excel");
        response.setCharacterEncoding("utf-8");
        String fileName = URLEncoder.encode("合并_" + System.currentTimeMillis(), "UTF-8");
        response.setHeader("Content-disposition", "attachment;filename=" + fileName + ".xlsx");
        mergeExcel(response);
    }

    public synchronized void mergeExcel(HttpServletResponse response) throws Exception {
        // 项目根目录
        String proFilePath = System.getProperty("user.dir");
        String dirPath = proFilePath + File.separator + "temp";
        File dir = new File(dirPath);
        if (!dir.exists()) {// 创建临时目录
            dir.mkdir();
        }
        File[] files = dir.listFiles();
        if (files != null && files.length > 0) {
            List<File> excelFiles = new ArrayList<>();
            for (File file : files) {
                if (file.getName().toLowerCase().endsWith(".xlsx") || file.getName().toLowerCase().endsWith(".xls")) {
                    excelFiles.add(file);
                }
            }
            if (excelFiles.size() > 0) {
                if (excelFiles.size() == 1) {
                    ExcelUtil.copyExcel(new FileInputStream(excelFiles.get(0)), response.getOutputStream(), "xlsx");
                } else {
                    // excel，有两行表头
                    ExcelUtil.mergeToOneSheet(response.getOutputStream(), excelFiles, 2);
//                    ExcelUtil.mergeToMultipleSheet(response.getOutputStream(), excelFiles);
                }
            }
        }
    }
}
