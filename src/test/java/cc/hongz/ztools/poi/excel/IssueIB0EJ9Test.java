package cc.hongz.ztools.poi.excel;

import cc.hongz.ztools.core.io.FileUtil;
import cc.hongz.ztools.core.lang.Console;
import org.junit.jupiter.api.Disabled;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

@DisplayName("issue IB0EJ9：Sax 方式读取数值型 Excel")
public class IssueIB0EJ9Test {
    @Test
    @Disabled
    @DisplayName("Sax 方式读取数值型 Excel 并打印行数据")
    void saxReadTest() {
        ExcelUtil.readBySax(FileUtil.file("d:/test/数值型测试.xlsx"), "hcm工资表",
            (sheetIndex, rowIndex, rowlist) -> Console.log("[{}] [{}] {}", sheetIndex, rowIndex, rowlist));
    }
}
