package cc.hongz.ztools.poi.word;

import cc.hongz.ztools.core.io.FileUtil;
import cc.hongz.ztools.core.io.IORuntimeException;
import cc.hongz.ztools.poi.exceptions.POIException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.File;
import java.io.IOException;

/**
 * Word Document工具。
 *
 * @author looly
 * @since 4.4.1
 */
@SuppressWarnings("checkstyle:HideUtilityClassConstructor")
public class DocUtil {

    /**
     * 创建{@link XWPFDocument}，如果文件已存在则读取之，否则创建新的。
     *
     * @param file docx文件
     * @return {@link XWPFDocument}
     */
    public static XWPFDocument create(final File file) {
        try {
            return FileUtil.exist(file) ? new XWPFDocument(OPCPackage.open(file)) : new XWPFDocument();
        } catch (InvalidFormatException e) {
            throw new POIException(e);
        } catch (IOException e) {
            throw new IORuntimeException(e);
        }
    }
}
