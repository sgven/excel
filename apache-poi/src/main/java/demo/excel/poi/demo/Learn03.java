package demo.excel.poi.demo;

import org.apache.poi.hssf.extractor.ExcelExtractor;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;

/**
 * 获取单元格内容
 * 文本提取
 * file vs inputStream
 */
public class Learn03 {

    public static void main(String[] args) throws Exception {
        Learn03 learn03 = new Learn03();
        learn03.getCellContents();
        learn03.textExtraction();
        learn03.fileVsInputStream();
        learn03.fileVsInputStream2();
    }

    /**
     * 获取单元格内容 Getting the cell contents
     */
    private void getCellContents() throws Exception {
        Workbook wb = new HSSFWorkbook();
        DataFormatter formatter = new DataFormatter();
        Sheet sheet1 = wb.getSheetAt(0);
        for (Row row : sheet1) {
            for (Cell cell : row) {
                CellReference cellRef = new CellReference(row.getRowNum(), cell.getColumnIndex());
                System.out.print(cellRef.formatAsString());
                System.out.print(" - ");
                // get the text that appears in the cell by getting the cell value and applying any data formats (Date, 0.00, 1.23e9, $1.23, etc)
                String text = formatter.formatCellValue(cell);
                System.out.println(text);
                // Alternatively, get the value and format it yourself
                switch (cell.getCellTypeEnum()) {
                    case STRING:
                        System.out.println(cell.getRichStringCellValue().getString());
                        break;
                    case NUMERIC:
                        if (DateUtil.isCellDateFormatted(cell)) {
                            System.out.println(cell.getDateCellValue());
                        } else {
                            System.out.println(cell.getNumericCellValue());
                        }
                        break;
                    case BOOLEAN:
                        System.out.println(cell.getBooleanCellValue());
                        break;
                    case FORMULA:
                        System.out.println(cell.getCellFormula());
                        break;
                    case BLANK:
                        System.out.println();
                        break;
                    default:
                        System.out.println();
                }
            }
        }
    }

    /**
     * 文本提取 Text Extraction
     * 对于大多数文本提取需求，标准的ExcelExtractor类应该会提供您所需要的全部。
     */
    private void textExtraction() throws Exception {
        try (InputStream inp = new FileInputStream("workbook.xls")) {
            HSSFWorkbook wb = new HSSFWorkbook(new POIFSFileSystem(inp));
            ExcelExtractor extractor = new ExcelExtractor(wb);
            extractor.setFormulasNotResults(true);
            extractor.setIncludeSheetNames(false);
            String text = extractor.getText();
            wb.close();
        }
    }

    /**
     * Files vs InputStreams
     * 1.使用WorkbookFactory，非常容易
     */
    private void fileVsInputStream() throws Exception {
        // Use a file
        Workbook wb = WorkbookFactory.create(new File("MyExcel.xls"));
        // Use an InputStream, needs more memory
        Workbook wb2 = WorkbookFactory.create(new FileInputStream("MyExcel.xlsx"));
    }

    /**
     * Files vs InputStreams
     * 2.直接使用 HSSFWorkbook or XSSFWorkbook
     * 通常应通过POIFSFileSystem或 OPCPackage来完全控制生命周期（包括完成后关闭文件）
     */
    private void fileVsInputStream2() throws Exception {
        InputStream myInputStream = null;
        // HSSFWorkbook, File
        POIFSFileSystem fs = new POIFSFileSystem(new File("file.xls"));
        HSSFWorkbook wb = new HSSFWorkbook(fs.getRoot(), true);
        // ...
        fs.close();
        // HSSFWorkbook, InputStream, needs more memory
        POIFSFileSystem fs2 = new POIFSFileSystem(myInputStream);
        HSSFWorkbook wb2 = new HSSFWorkbook(fs.getRoot(), true);

        // XSSFWorkbook, File
        OPCPackage pkg = OPCPackage.open(new File("file.xlsx"));
        XSSFWorkbook wb3 = new XSSFWorkbook(pkg);
        // ...
        pkg.close();
        // XSSFWorkbook, InputStream, needs more memory
        OPCPackage pkg2 = OPCPackage.open(myInputStream);
        XSSFWorkbook wb4 = new XSSFWorkbook(pkg);
        // ...
        pkg.close();
    }

}
