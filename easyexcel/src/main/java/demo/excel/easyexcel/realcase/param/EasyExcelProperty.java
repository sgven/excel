package demo.excel.easyexcel.realcase.param;

/**
 * 导出信息配置
 */
public class EasyExcelProperty {

    private String fileName; // 文件名，必传
    private String titleName; // 标题，若为空，则取fileName
    private String sheetName;
    private Class clazz; // 必传

    public String getFileName() {
        return fileName;
    }

    public void setFileName(String fileName) {
        this.fileName = fileName;
    }

    public String getTitleName() {
        return titleName;
    }

    public void setTitleName(String titleName) {
        this.titleName = titleName;
    }

    public String getSheetName() {
        return sheetName;
    }

    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }

    public Class getClazz() {
        return clazz;
    }

    public void setClazz(Class clazz) {
        this.clazz = clazz;
    }

}
