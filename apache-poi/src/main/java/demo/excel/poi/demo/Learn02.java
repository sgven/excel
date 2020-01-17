package demo.excel.poi.demo;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * 遍历单元格
 */
public class Learn02 {

    private static final int MY_MINIMUM_COLUMN_COUNT = 0;

    public static void main(String[] args) {
        Learn02 learn02 = new Learn02();
        learn02.loop();
        learn02.loopAllCell();
    }

    /**
     * Iterate over rows and cells
     * 想遍历所有workbook中的sheet，sheet中的所有row，row中的所有cell。有以下两种方式：
     *
     * 1.通过调用下面的方法，使用迭代器循环。
     * 注意，rowIterator和cellIterator遍历已创建的行或单元格，跳过空的行和单元格。
     * wb.sheetIterator();
     * sheet.rowIterator();
     * row.cellIterator();
     *
     * 2.or 隐式使用for-each loop
     */
    private void loop() {
        Workbook wb = new HSSFWorkbook();
        for (Sheet sheet : wb) {
            for (Row row : sheet) {
                for (Cell cell : row) {
                    // Do something here
                }
            }
        }
    }

    /**
     * Iterate over cells, with control of missing / blank cells
     * 某些情况下，循环迭代时，你需要完全控制如何处理丢失或空白的行和单元格，
     * 并且需要确保访问每个单元格，而不仅仅是访问文件中定义的那些单元格。
     * （CellIterator将仅返回文件中定义的单元格（有值或有样式），取决于Excel）
     *
     * 调用getCell（int，MissingCellPolicy） 来获取单元格，使用 MissingCellPolicy 控制空白或空单元格的处理方式。
     */
    private void loopAllCell() {
        Workbook wb = new HSSFWorkbook();
        Sheet sheet = wb.getSheet("0");
        // Decide which rows to process
        int rowStart = Math.min(15, sheet.getFirstRowNum());
        int rowEnd = Math.max(1400, sheet.getLastRowNum());
        for (int rowNum = rowStart; rowNum < rowEnd; rowNum++) {
            Row r = sheet.getRow(rowNum);
            if (r == null) {
                // This whole row is empty
                // Handle it as needed
                continue;
            }
            int lastColumn = Math.max(r.getLastCellNum(), MY_MINIMUM_COLUMN_COUNT);
            for (int cn = 0; cn < lastColumn; cn++) {
                Cell c = r.getCell(cn, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                if (c == null) {
                    // The spreadsheet is empty in this cell（对excel中的空单元格处理）
                } else {
                    // Do something useful with the cell's contents（对单元格内容进行一些有用的操作）
                }
            }
        }
    }

}
