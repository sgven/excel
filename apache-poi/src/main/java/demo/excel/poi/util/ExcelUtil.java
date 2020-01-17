package demo.excel.poi.util;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.util.StringUtils;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.Future;
import java.util.concurrent.TimeUnit;

public class ExcelUtil {

    private static final Logger logger = LoggerFactory.getLogger(ExcelUtil.class);

    /**
     * 复制excel
     * @param inputStream
     * @param outputStream
     * @param type  excel类型，xlsx 和 xls
     * @throws Exception
     */
    public static void copyExcel(InputStream inputStream, OutputStream outputStream, String type) throws Exception {
        if (StringUtils.isEmpty(type)) {
            throw new RuntimeException("excel类型不能为空");
        }
        Workbook workbook;
        if (type.toLowerCase().equals("xlsx")) {
            workbook = new XSSFWorkbook(inputStream);
        } else {
            workbook = new HSSFWorkbook(inputStream);
        }
        workbook.write(outputStream);
    }

    /**
     * 对同一种列表形式数据的多个excel文件，合并到一个excel的同一个sheet中
     * @param outputStream
     * @param files
     * @param beginMergeRowIndex
     * @throws Exception
     */
    public static void mergeToOneSheet(OutputStream outputStream, List<File> files, int beginMergeRowIndex) throws Exception {
        if (files != null && files.size() > 0) {
            long readBeginTime = System.currentTimeMillis();
            ExecutorService executor = Executors.newFixedThreadPool(files.size());
            List<Future<Workbook>> readCache = new ArrayList<>();
            // 读（并发读）
            for (File file : files) {
                Future<Workbook> future = (Future<Workbook>) executor.submit(() -> {
                    return readFile(file);
                });
                readCache.add(future);
            }
            // 等待线程池任务执行完毕
            executor.shutdown();
            while (!executor.isTerminated()) {
                System.out.println("任务执行中");
//                executor.awaitTermination(100, TimeUnit.MILLISECONDS);
                executor.awaitTermination(1, TimeUnit.SECONDS);
            }
            long readEndTime = System.currentTimeMillis();
            System.out.println("读取耗时：" + (readEndTime - readBeginTime));
            // 写（顺序写）
            // 创建Workbook，暂存数据
            Workbook targetWorkbook = new XSSFWorkbook();
            int targetRowIndex = 0;// 记录写入行的位置
            boolean isFirstSheet = true;// 是否是第一个sheet
            copyFile2OneSheet(readCache, targetWorkbook, beginMergeRowIndex, targetRowIndex, isFirstSheet);
            System.out.println("写总耗时：" + (System.currentTimeMillis() - readEndTime));
            targetWorkbook.write(outputStream);
            System.out.println("合并总耗时：" + (System.currentTimeMillis() - readBeginTime));
        }
    }

    /**
     * 将多个excel文件，合并到一个excel的多个sheet中
     * @param outputStream
     * @param files
     * @throws Exception
     */
    public static void mergeToMultipleSheet(OutputStream outputStream, List<File> files) throws Exception {
        if (files != null && files.size() > 0) {
            long readBeginTime = System.currentTimeMillis();
            ExecutorService executor = Executors.newFixedThreadPool(files.size());
            List<Future<Workbook>> readCache = new ArrayList<>();
            // 读（并发读）
            for (File file : files) {
                Future<Workbook> future = (Future<Workbook>) executor.submit(() -> {
                    return readFile(file);
                });
                readCache.add(future);
            }
            // 等待线程池任务执行完毕
            executor.shutdown();
            while (!executor.isTerminated()) {
                System.out.println("任务执行中");
//                executor.awaitTermination(100, TimeUnit.MILLISECONDS);
                executor.awaitTermination(1, TimeUnit.SECONDS);
            }
            long readEndTime = System.currentTimeMillis();
            System.out.println("读取耗时：" + (readEndTime - readBeginTime));
            // 写（顺序写）
            // 创建Workbook，暂存数据
            Workbook targetWorkbook = new XSSFWorkbook();
            int targetRowIndex = 0;// 记录写入行的位置
            copyFile2MultipleSheet(readCache, targetWorkbook, targetRowIndex);
            System.out.println("写总耗时：" + (System.currentTimeMillis() - readEndTime));
            targetWorkbook.write(outputStream);
            System.out.println("合并总耗时：" + (System.currentTimeMillis() - readBeginTime));
        }
    }

    private static void copyFile2MultipleSheet(List<Future<Workbook>> readCache, Workbook targetWorkbook, int targetRowIndex) throws Exception {
        for (Future<Workbook> future : readCache) {
            Workbook sourceWorkbook = future.get();
            int sheetNum = sourceWorkbook.getNumberOfSheets();
            if (sheetNum == 0) {
                continue;
            }
            for (int i = 0; i < sheetNum; i++) {
                Sheet sourceSheet = sourceWorkbook.getSheetAt(i);
                Sheet targetSheet = targetWorkbook.createSheet();
                copySheet(targetWorkbook, sourceSheet, targetSheet, 0, targetRowIndex);
            }
        }
    }

    private static void copyFile2OneSheet(List<Future<Workbook>> readCache, Workbook targetWorkbook,
                                                       int beginMergeRowIndex, int targetRowIndex, boolean isFirstSheet) throws InterruptedException, java.util.concurrent.ExecutionException {
        Sheet targetSheet = targetWorkbook.createSheet();
        for (Future<Workbook> future : readCache) {
            Workbook sourceWorkbook = future.get();
            int sheetNum = sourceWorkbook.getNumberOfSheets();
            if (sheetNum == 0) {
                continue;
            }
            long beginTime = System.currentTimeMillis();
            for (int i = 0; i < sheetNum; i++) {
                Sheet sourceSheet = sourceWorkbook.getSheetAt(i);
                if (isFirstSheet) {// 取得第一个sheet的表头
                    int titleRowSize = beginMergeRowIndex;
                    for (int j = 0; j < titleRowSize; j++) {
                        Row sourceRow = sourceSheet.getRow(j);
                        Row targetRow = targetSheet.createRow(targetRowIndex++);
                        copyRow(targetWorkbook, sourceRow, targetRow);
                    }
                    isFirstSheet = false;
                }
                targetRowIndex = targetRowIndex + copySheet(targetWorkbook, sourceSheet, targetSheet, beginMergeRowIndex, targetRowIndex);
            }
            System.out.println("已写行数：" + targetRowIndex + "，耗时：" + (System.currentTimeMillis() - beginTime));
        }
    }

    private static Workbook readFile(File file) throws Exception {
        if (file == null || !file.getName().toLowerCase().endsWith(".xlsx") &&
                !file.getName().toLowerCase().endsWith(".xls")) {
            logger.error("文件格式错误");
            throw new RuntimeException("文件格式错误");
        }
        return WorkbookFactory.create(new FileInputStream(file));
    }

    /**
     *
     * @param wb
     * @param srcSheet
     * @param targetSheet
     * @param srcBeginRowIndex
     * @param targetRowIndex
     * @return 复制的行数
     */
    private static int copySheet(Workbook wb, Sheet srcSheet, Sheet targetSheet, int srcBeginRowIndex, int targetRowIndex) {
//        mergeSheetAllRegion(srcSheet, targetSheet);
        //设置列宽
        /*for(int i = srcBeginRowIndex; i <= srcSheet.getRow(srcSheet.getFirstRowNum()).getLastCellNum(); i++){
            targetSheet.setColumnWidth(i,srcSheet.getColumnWidth(i));
        }*/
        int i = srcBeginRowIndex;
        for (;i <= srcSheet.getLastRowNum(); i ++) {// 用<=，是因为平台给的excel文件，实际上比这里最后一行多一行
            Row srcRow = srcSheet.getRow(i);
            if (srcRow == null) {
                continue;
            }
            Row targetRow = targetSheet.createRow(targetRowIndex++);
            copyRow(wb, srcRow, targetRow);
        }
        return i-srcBeginRowIndex;
    }

    private static void mergeSheetAllRegion(Sheet srcSheet, Sheet targetSheet) {//合并单元格
        int num = srcSheet.getNumMergedRegions();
        CellRangeAddress cellR = null;
        for (int i = 0; i < num; i++) {
            cellR = srcSheet.getMergedRegion(i);
            targetSheet.addMergedRegion(cellR);
        }
    }

    private static void copyRow(Workbook wb, Row srcRow, Row targetRow){
//        targetRow.setHeight(srcRow.getHeight());
        for (int j = 0; j < srcRow.getLastCellNum(); j++) {//Cell
            Cell sourceCell = srcRow.getCell(j);
            if (sourceCell == null) {
                continue;
            }
            Cell targetCell = targetRow.createCell(j);
            copyCell(wb, sourceCell, targetCell);
        }
    }

    private static void copyCell(Workbook wb, Cell srcCell, Cell targetCell) {
        //targetCell.setEncoding(srcCell.getEncoding());
        /*//样式
        CellStyle targetStyle= wb.createCellStyle();
        copyCellStyle(srcCell.getCellStyle(), targetStyle);
        targetCell.setCellStyle(targetStyle);*/
        if (srcCell.getCellComment() != null) {
            targetCell.setCellComment(srcCell.getCellComment());
        }
        // 不同数据类型处理
        int fromCellType = srcCell.getCellType();
        targetCell.setCellType(fromCellType);
        if (fromCellType == Cell.CELL_TYPE_NUMERIC) {
            if (DateUtil.isCellDateFormatted(srcCell)) {
                targetCell.setCellValue(srcCell.getDateCellValue());
            } else {
                targetCell.setCellValue(srcCell.getNumericCellValue());
            }
        } else if (fromCellType == Cell.CELL_TYPE_STRING) {
            targetCell.setCellValue(srcCell.getRichStringCellValue());
        } else if (fromCellType == Cell.CELL_TYPE_BLANK) {
            // nothing21
        } else if (fromCellType == Cell.CELL_TYPE_BOOLEAN) {
            targetCell.setCellValue(srcCell.getBooleanCellValue());
        } else if (fromCellType == Cell.CELL_TYPE_ERROR) {
            targetCell.setCellErrorValue(srcCell.getErrorCellValue());
        } else if (fromCellType == Cell.CELL_TYPE_FORMULA) {
            targetCell.setCellFormula(srcCell.getCellFormula());
        } else { // nothing29
        }

    }

    private static void copyCellStyle(CellStyle srcStyle, CellStyle targetStyle) {
        targetStyle.cloneStyleFrom(srcStyle);
    }
}
