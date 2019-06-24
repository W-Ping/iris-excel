package com.iris.excelfile.utils;

import com.iris.excelfile.exception.ExcelParseException;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.openxml4j.opc.OPCPackage;

import java.io.*;
import java.net.URL;

@Slf4j
public class FileUtil {
    private static final String JAVA_IO_TMPDIR = "java.io.tmpdir";
    private static final String FILETEMP = "fileTemp";

    public static InputStream getResourcesFileInputStream(String fileName) {
        return Thread.currentThread().getContextClassLoader().getResourceAsStream("" + fileName);
    }

    public static synchronized InputStream synGetResourcesFileInputStream(String fileName) throws IOException {
        return new FileInputStream(fileName);
    }

    /**
     * @param filePath
     * @param fileName
     * @return
     * @throws IOException
     */
    public static synchronized OutputStream synGetResourcesFileOutputStream(String filePath, String fileName) throws IOException {
        File file = new File(fileReplacePath(filePath));
        if (!file.exists()) {
            file.mkdirs();
        }
        file = new File(fileReplacePath(filePath + "/" + fileName));
        return new FileOutputStream(file);
    }

    /**
     * @param filePath
     * @return
     */
    public static String fileReplacePath(String filePath) {
        String os = System.getProperty("os.name");
        if (os.toLowerCase().indexOf("windows") != -1) {
            filePath = filePath.replace("/", "\\");
        } else {
            filePath = filePath.replace("\\", "/");
        }
        return filePath;
    }

    /**
     * @param inputStream
     * @param outputStream
     */
    public static void close(InputStream inputStream, OutputStream outputStream) {
        try {
            if (inputStream != null) {
                inputStream.close();
            }
            if (outputStream != null) {
                outputStream.close();
            }
        } catch (IOException ex) {
            log.error("file IO close error {}", ex.getMessage());
            throw new ExcelParseException("file IO close error");
        }
    }

    public static void close(OPCPackage opc) {
        try {
            if (opc != null) {
                opc.close();
            }
        } catch (Exception ex) {
            log.error("OPCPackage close error {}", ex.getMessage());
            throw new ExcelParseException("OPCPackage close error");
        }
    }

    /**
     * @param filePath
     * @return
     */
    public static boolean checkFilePath(String filePath) {
        File file = new File(fileReplacePath(filePath));
        if (!file.exists()) {
            throw new ExcelParseException("系统找不到指定的文件路径【" + filePath + "】");
        }
        return true;
    }

    /**
     * @return
     */
    public static String createTempDirectory() {

        String tmpDir = System.getProperty(JAVA_IO_TMPDIR);
        if (tmpDir == null) {
            throw new ExcelParseException(
                    "Systems temporary directory not defined - set the -D" + JAVA_IO_TMPDIR + " jvm fileTemp!");
        }
        File directory = new File(tmpDir, FILETEMP);
        if (!directory.exists()) {
            syncCreateTempFilesDirectory(directory);
        }
        return directory.getAbsolutePath();
    }

    /**
     * @param directory
     */
    private static synchronized void syncCreateTempFilesDirectory(File directory) {
        if (!directory.exists()) {
            directory.mkdirs();
        }
    }
}
