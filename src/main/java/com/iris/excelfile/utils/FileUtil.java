package com.iris.excelfile.utils;

import com.iris.excelfile.constant.ExcelTypeEnum;
import com.iris.excelfile.exception.ExcelParseException;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.openxml4j.opc.OPCPackage;

import java.io.*;
import java.time.LocalDate;
import java.util.UUID;

/**
 *
 * @author lwp
 */
@Slf4j
public class FileUtil {
	private static final String JAVA_IO_TMPDIR = "java.io.tmpdir";
	private static final String FILE_TEMP = "exportFileTemp";

	public static InputStream getResourcesFileInputStream(String fileName) {
		return Thread.currentThread().getContextClassLoader().getResourceAsStream("" + fileName);
	}

	public static synchronized InputStream synGetResourcesFileInputStream(String fileName) throws IOException {
		return new FileInputStream(fileName);
	}

	/**
	 * @param filePath
	 * @return
	 * @throws IOException
	 */
	public static synchronized void synMkdirs(String filePath) {
		File file = new File(fileReplacePath(filePath));
		if (!file.exists()) {
			file.mkdirs();
		}
	}

	public static synchronized OutputStream synGetResourcesFileOutputStream(String excelOutFileFullPath) throws IOException {
		String fileDir = getFileDir(excelOutFileFullPath);
		if (fileDir != null) {
			synMkdirs(fileDir);
		}
		File file = new File(fileReplacePath(excelOutFileFullPath));
		if (!file.exists() && file.isFile()) {
			file.createNewFile();
		}
		return new FileOutputStream(file);
	}

	/**
	 * @param filePath
	 * @return
	 */
	public static String fileReplacePath(String filePath) {
		if (isWindowSystem()) {
			filePath = filePath.replace("/", "\\");
		} else {
			filePath = filePath.replace("\\", "/");
		}
		return filePath;
	}

	/**
	 * @param filePath
	 * @return
	 */
	public static String getFileDir(String filePath) {
		if (StringUtils.isNotBlank(filePath)) {
			String pathSplit = "\\";
			if (!isWindowSystem()) {
				pathSplit = "/";
			}
			if (filePath.indexOf(pathSplit) >= 0) {
				return filePath.substring(0, filePath.lastIndexOf(pathSplit));
			}
		}
		return null;
	}

	/**
	 * @param filePath
	 * @return
	 */
	public static String getFileName(String filePath) {
		if (StringUtils.isNotBlank(filePath)) {
			String pathSplit = "\\";
			if (!isWindowSystem()) {
				pathSplit = "/";
			}
			if (filePath.indexOf(pathSplit) >= 0) {
				String fileName = filePath.substring(filePath.lastIndexOf(pathSplit) + 1, filePath.length());
				return fileName;
			}
		}
		return null;
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
		File directory = new File(tmpDir, FILE_TEMP);
		if (!directory.exists()) {
			syncCreateTempFilesDirectory(directory);
		}
		return directory.getAbsolutePath();
	}

	/**
	 * @param path
	 * @param excelTypeEnum
	 * @return
	 */
	public static synchronized String createTempOutFile(String path, ExcelTypeEnum excelTypeEnum) {
		File file = new File(fileReplacePath(path + "/" + UUID.randomUUID().toString() + "_" + LocalDate.now() + excelTypeEnum.getValue()));
		if (!file.isFile()) {
			try {
				file.createNewFile();
			} catch (Exception e) {
				e.printStackTrace();
			}
		}
		return file.getPath();
	}

	/**
	 * @param directory
	 */
	private static synchronized void syncCreateTempFilesDirectory(File directory) {
		if (!directory.exists()) {
			directory.mkdirs();
		}
	}

	/**
	 * 是否是Window系统
	 *
	 * @return
	 */
	private static boolean isWindowSystem() {
		return System.getProperty("os.name").toLowerCase().indexOf("windows") >= 0;
	}
}
