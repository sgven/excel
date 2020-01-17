package demo.excel.poi.demo;

import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;

/**
 * 读写
 */
public class Learn05 {

    public static void main(String[] args) throws Exception {
        Learn05 learn05 = new Learn05();
        learn05.readAndReWriting();
    }

    /**
     * 读取和重写工作簿 Reading and Rewriting Workbooks
     */
    private void readAndReWriting() throws Exception {
        try (InputStream inp = new FileInputStream("workbook.xls")) {
            //InputStream inp = new FileInputStream("workbook.xlsx");
            Workbook wb = WorkbookFactory.create(inp);
            Sheet sheet = wb.getSheetAt(0);
            Row row = sheet.getRow(2);
            Cell cell = row.getCell(3);
            if (cell == null)
                cell = row.createCell(3);
            cell.setCellType(CellType.STRING);
            cell.setCellValue("a test");
            // Write the output to a file
            try (OutputStream fileOut = new FileOutputStream("workbook.xls")) {
                wb.write(fileOut);
            }
        }
    }

}
