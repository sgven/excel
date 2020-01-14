package demo.excel.easyexcel.realcase.handler;

import com.alibaba.excel.annotation.ExcelIgnore;
import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.metadata.CellData;
import com.alibaba.excel.metadata.Head;
import com.alibaba.excel.write.handler.CellWriteHandler;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.metadata.holder.WriteTableHolder;
import org.apache.poi.ss.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

/**
 * 自定义拦截器
 *
 */
public class CustomCellWriteHandler implements CellWriteHandler {

    private final Logger LOGGER = LoggerFactory.getLogger(CustomCellWriteHandler.class);

    private Class clazz;

    public CustomCellWriteHandler() {
    }

    public CustomCellWriteHandler(Class clazz) {
        this.clazz = clazz;
    }

    @Override
    public void beforeCellCreate(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, Row row,
                                 Head head, Integer columnIndex, Integer relativeRowIndex, Boolean isHead) {

    }

    @Override
    public void afterCellCreate(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, Cell cell,
                                Head head, Integer relativeRowIndex, Boolean isHead) {

    }

    @Override
    public void afterCellDispose(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder,
                                 List<CellData> cellDataList, Cell cell, Head head, Integer relativeRowIndex, Boolean isHead) {
        // 这里可以对cell进行任何操作
        LOGGER.info("第{}行，第{}列写入完成。", cell.getRowIndex(), cell.getColumnIndex());
        // 设置样式
        setStyle(writeSheetHolder, cell, isHead);
    }

    private void setStyle(WriteSheetHolder writeSheetHolder, Cell cell, Boolean isHead) {
        Workbook workbook = writeSheetHolder.getSheet().getWorkbook();
        if (isHead) {
            // 设置第一行标题栏的样式
            if (cell.getRowIndex() == 0) {
                cell.setCellStyle(getTitleStyle(workbook));
                cell.getRow().setHeight(((short) 500));
            }
            // 设置第二行表头列的样式
            if (cell.getRowIndex() == 1) {
                cell.setCellStyle(getHeaderStyle(workbook));
            }
        } else {
            cell.setCellStyle(getBodyStyle(workbook));
            /*
            // ExcelExportUtil的设置列宽的逻辑
            Sheet sheet = writeSheetHolder.getSheet();
            List<String> columns = getHeadColumnList(clazz);
            for (int i = 0; i < columns.size(); i ++) {
                String column = columns.get(i);
                if (column.getBytes().length > 100) {
                    sheet.setColumnWidth(i, 17920);
                } else {
                    sheet.setColumnWidth(i, column.getBytes().length * 384);
                }
            }*/
        }
    }

    private List<String> getHeadColumnList(Class clazz) {
        List<String> columns = new ArrayList<>();
        Field[] fields = clazz.getDeclaredFields();
        if (fields != null && fields.length > 0) {
            Arrays.stream(fields).forEach(field -> {
                ExcelProperty columnProperty = field.getAnnotation(ExcelProperty.class);
                ExcelIgnore columnIgnore = field.getAnnotation(ExcelIgnore.class);
                if (columnIgnore == null && columnProperty != null) {
                    columns.add(columnProperty.value()[0]);
                }
            });
        }
        return columns;
    }

    private CellStyle getTitleStyle(Workbook workbook) {
        CellStyle cellStyle = setCellStyle(workbook);
        Font font = workbook.createFont();
        font.setBold(true);
        font.setFontHeightInPoints((short) 18);
        cellStyle.setFont(font);
        return cellStyle;
    }

    private static CellStyle getHeaderStyle(Workbook workbook) {
        CellStyle cellStyle = setCellStyle(workbook);
        Font font = workbook.createFont();
        font.setBold(true);
        font.setFontHeightInPoints((short) 12);
        cellStyle.setFont(font);
        return cellStyle;
    }

    private static CellStyle getBodyStyle(Workbook book) {
        CellStyle cellStyle = setCellStyle(book);
        cellStyle.setWrapText(true);
        return cellStyle;
    }

    private static CellStyle setCellStyle(Workbook workbook) {
        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setBorderTop(BorderStyle.THIN);
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBorderLeft(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setFillBackgroundColor(IndexedColors.WHITE.getIndex());
        return cellStyle;
    }
}
