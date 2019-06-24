package com.iris.excelfile.utils;

import com.iris.excelfile.constant.ExcelTypeEnum;
import com.iris.excelfile.exception.ExcelParseException;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.poifs.crypt.Decryptor;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.crypt.EncryptionMode;
import org.apache.poi.poifs.crypt.Encryptor;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.*;
import java.security.GeneralSecurityException;

/**
 *
 */
@Slf4j
public class ExcelV2007Util {
    /**
     * 加密
     *
     * @param outFilePath
     * @param password
     * @return
     */
    public static boolean encrypt(String outFilePath, String password) {
        if (StringUtils.isBlank(password)) {
            return false;
        }
        FileOutputStream fos = null;
        OPCPackage opc = null;
        try {
            POIFSFileSystem fs = new POIFSFileSystem();
            EncryptionInfo info = new EncryptionInfo(EncryptionMode.standard);
            Encryptor enc = info.getEncryptor();
            //设置密码
            enc.confirmPassword(password);
            //加密文件
            opc = OPCPackage.open(new File(outFilePath), PackageAccess.READ_WRITE);
            OutputStream os = enc.getDataStream(fs);
            opc.save(os);
            FileUtil.close(opc);
            //把加密后的文件写回到流
            fos = new FileOutputStream(outFilePath);
            fs.writeFilesystem(fos);
        } catch (Exception e) {
            log.error("文档加密失败【{}】", outFilePath);
            throw new ExcelParseException("文档加解密失败");
        } finally {
            FileUtil.close(null, fos);
        }
        return true;
    }

    /**
     * 解密
     *
     * @param filePath
     * @param password
     * @return
     */
    public static Workbook decrypt(String filePath, String password) {
        Workbook wb = null;
        FileInputStream in = null;
        try {
            in = new FileInputStream(filePath);
            POIFSFileSystem poifsFileSystem = new POIFSFileSystem(in);
            EncryptionInfo encInfo = new EncryptionInfo(poifsFileSystem);
            Decryptor decryptor = Decryptor.getInstance(encInfo);
            if (!decryptor.verifyPassword(password)) {
                throw new ExcelParseException("无法处理，解密文档失败！");
            }
            wb = WorkBookUtil.createWorkBook(decryptor.getDataStream(poifsFileSystem), ExcelTypeEnum.XLSX);
            //Workbook wb = WorkbookFactory.create(in, password);
        } catch (IOException e) {
            log.error("处理加密文档I/O异常【{}】", filePath);
            throw new ExcelParseException("无法处理加密文档");
        } catch (GeneralSecurityException e) {
            log.error("无法处理加密文档【{}】", filePath);
            throw new ExcelParseException("无法处理加密文档");
        } finally {
            FileUtil.close(in, null);
        }
        return wb;
    }
}
