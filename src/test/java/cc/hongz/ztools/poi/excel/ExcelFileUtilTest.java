package cc.hongz.ztools.poi.excel;

import cc.hongz.ztools.core.io.FileUtil;
import cc.hongz.ztools.core.io.IoUtil;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import java.io.InputStream;

@DisplayName("Excel文件类型判断测试")
public class ExcelFileUtilTest {

    @DisplayName("xls文件类型判断测试")
    @Test
    public void xlsTest() {
        InputStream in = FileUtil.getInputStream("aaa.xls");
        try {
            assertTrue(ExcelFileUtil.isXls(in));
            assertFalse(ExcelFileUtil.isXlsx(in));
        } finally {
            IoUtil.close(in);
        }
    }

    @DisplayName("xlsx文件类型判断测试")
    @Test
    public void xlsxTest() {
        InputStream in = FileUtil.getInputStream("aaa.xlsx");
        try {
            assertFalse(ExcelFileUtil.isXls(in));
            assertTrue(ExcelFileUtil.isXlsx(in));
        } finally {
            IoUtil.close(in);
        }
    }
}
