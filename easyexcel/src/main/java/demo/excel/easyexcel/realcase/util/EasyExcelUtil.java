package demo.excel.easyexcel.realcase.util;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelReader;
import com.alibaba.excel.analysis.ExcelReadExecutor;
import com.alibaba.excel.read.builder.ExcelReaderBuilder;
import com.alibaba.excel.read.listener.ReadListener;
import com.alibaba.excel.read.metadata.ReadSheet;
import demo.excel.easyexcel.demo.read.ReadDemoData;
import demo.excel.easyexcel.realcase.listener.AbstractReadListener;
import demo.excel.easyexcel.realcase.listener.NoModleDataListener;

import java.io.*;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

public class EasyExcelUtil {

    /**
     * 对同一种列表形式数据的多个excel文件，合并到一个excel的同一个sheet中
     * @param outputStream
     * @param files
     * @param skipRowCount  除第一个文件外，后合并的文件需要跳过的行数
     * @throws Exception
     */
    public static void merge2OneSheet(OutputStream outputStream, List<File> files, int skipRowCount) throws Exception {
        // 读
        List<Map<Integer, String>> all = read(files, skipRowCount);
        // 写
        List<List<Object>> list = parse2List(all);
        // 注意，必须以这种方式写，否则垃圾回收会提前触发excelWriter.finish()
        EasyExcel.write(outputStream).sheet(0).doWrite(list);
    }

    private static List read(List<File> files, int skipRowCount) throws Exception {
        List<Map<Integer, String>> all = new ArrayList<>();
        for (int i = 0; i < files.size(); i++) {
            // 第一个文件不跳过表头
            int skipCount = i == 0 ? 0 : skipRowCount;
            NoModleDataListener readListener = new NoModleDataListener(skipCount);
            all.addAll(readAllSheet(files.get(i), readListener, null, 0));
        }
        return all;
    }

    /**
     * 读取全部sheet
     * @param file
     * @param readListener
     * @param head  表头class
     * @param headRowNumber 表头的行数
     * @return
     * @throws Exception
     */
    public static List readAllSheet(File file, AbstractReadListener readListener, Class head, int headRowNumber) throws Exception {
        if (file == null || !file.getName().toLowerCase().endsWith("xlsx") && !file.getName().toLowerCase().endsWith("xls")) {
            throw new RuntimeException("文件格式错误");
        }
        return readAllSheet(new FileInputStream(file), readListener, head, headRowNumber);
    }

    /**
     * 读取全部sheet
     * @param inputStream
     * @param readListener
     * @param head  表头class
     * @param headRowNumber 表头的行数
     * @return
     * @throws Exception
     */
    public static List readAllSheet(InputStream inputStream, AbstractReadListener readListener, Class head, int headRowNumber) throws Exception {
        ExcelReader reader = getReader(inputStream, readListener, head, headRowNumber);
        ExcelReadExecutor excelReadExecutor = reader.excelExecutor();
        List<ReadSheet> sheets = excelReadExecutor.sheetList();
        for (ReadSheet sheet : sheets) {
            reader.read(sheet);
        }
        reader.finish();
        return readListener.getDatas();
    }

    /**
     * 读取单个sheet
     * @param file
     * @param readListener
     * @param head  表头class
     * @param headRowNumber 表头的行数
     * @return
     * @throws Exception
     */
    public static List readSingleSheet(File file, AbstractReadListener readListener, Class head, int headRowNumber) throws Exception {
        if (file == null || !file.getName().toLowerCase().endsWith("xlsx") && !file.getName().toLowerCase().endsWith("xls")) {
            throw new RuntimeException("文件格式错误");
        }
        return readSingleSheet(new FileInputStream(file), readListener, head, headRowNumber);
    }

    /**
     * 读取单个sheet
     * @param inputStream
     * @param readListener
     * @param head  表头class
     * @param headRowNumber 表头的行数
     * @return
     * @throws Exception
     */
    public static List readSingleSheet(InputStream inputStream, AbstractReadListener readListener, Class head, int headRowNumber) throws Exception {
        ExcelReader reader = getReader(inputStream, readListener, head, headRowNumber);
        ReadSheet readSheet = EasyExcel.readSheet(0).build();
        reader.read(readSheet);
        reader.finish();
        return readListener.getDatas();
    }

    private static List<List<Object>> parse2List(List<Map<Integer, String>> mapList) {
        List<List<Object>> back = new ArrayList<List<Object>>();
        for (Map<Integer, String> map : mapList) {
            List<Object> row = new ArrayList<Object>();
            for (Map.Entry<Integer, String> entry : map.entrySet()) {
                row.add(entry.getValue());
            }
            back.add(row);
        }
        return back;
    }

    private static ExcelReader getReader(InputStream inputStream, ReadListener readListener, Class head, int headRowNumber) {
        ExcelReaderBuilder readerBuilder = EasyExcel.read(inputStream);
        if (readListener != null) {
            readerBuilder.registerReadListener(readListener);
        }
        if (head != null) {
            readerBuilder.head(head);
        }
        readerBuilder.headRowNumber(headRowNumber);
        return readerBuilder.build();
    }

}
