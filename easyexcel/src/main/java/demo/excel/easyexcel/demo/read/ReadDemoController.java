package demo.excel.easyexcel.demo.read;

import demo.excel.easyexcel.demo.util.TestFileUtil;
import demo.excel.easyexcel.realcase.listener.NoModleDataListener;
import demo.excel.easyexcel.realcase.util.EasyExcelUtil;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;

import java.io.File;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

@Controller
@RequestMapping("read")
public class ReadDemoController {

    @RequestMapping(value = "simpleRead")
    public void simpleRead() throws Exception {
        // 有个很重要的点 DemoDataListener 不能被spring管理，要每次读取excel都要new,然后里面用到spring可以构造方法传进去
//        EasyExcel.read(fileName, ReadDemoData.class, new ReadListener()).sheet().doRead();
        // 无表头
        String noheadFileName = TestFileUtil.getPath() + "demo" + File.separator + "no_head_list.xlsx";
        List list0 = EasyExcelUtil.readSingleSheet(new File(noheadFileName), new ReadListener(), ReadDemoData.class, 0);
        // 有一行表头
        String oneheadFileName = TestFileUtil.getPath() + "demo" + File.separator + "one_head_list.xlsx";
        List list1 = EasyExcelUtil.readSingleSheet(new File(oneheadFileName), new ReadListener(), ReadDemoData.class, 1);
        System.out.println("指定对象读0：" + list0.size());
        System.out.println("指定对象读1：" + list1.size());
    }

    @RequestMapping(value = "mergeRead")
    public void mergeRead() throws Exception {
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
                List<Map<Integer, String>> all = new ArrayList<>();
                for (int i = 0; i < excelFiles.size(); i++) {
                    // 第一个文件不跳过表头
                    int skipCount = i == 0 ? 0 : 2;
                    NoModleDataListener readListener = new NoModleDataListener(skipCount);
                    all.addAll(EasyExcelUtil.readAllSheet(excelFiles.get(i), readListener, null, 0));
                }
                System.out.println("读取行数：" + all.size());
            }
        }
    }

}
