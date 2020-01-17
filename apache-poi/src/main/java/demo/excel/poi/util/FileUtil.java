package demo.excel.poi.util;

import java.io.File;

public class FileUtil {


    /**
     * 递归删除文件夹及文件
     * @param file
     * @return
     */
    public static boolean delDirAndFile(File file) {
        if (file == null || !file.exists()) {
            return false;
        }
        if (file.isDirectory()) {
            File[] files = file.listFiles();
            for (File f : files) {
                delFile(f);
            }
        }
        return file.delete();
    }

    /**
     * 递归删除文件夹下的文件
     * @param file
     * @return
     */
    public static boolean delFile(File file) {
        if (file == null || !file.exists()) {
            return false;
        }
        if (file.isDirectory()) {
            File[] files = file.listFiles();
            for (File f : files) {
                delFile(f);
            }
        } else {
            file.delete();
        }
        return true;
    }
}
