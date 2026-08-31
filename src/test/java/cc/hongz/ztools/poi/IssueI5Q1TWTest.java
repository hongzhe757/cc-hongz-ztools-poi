package cc.hongz.ztools.poi;

import cc.hongz.ztools.poi.excel.ExcelReader;
import cc.hongz.ztools.poi.excel.ExcelUtil;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.assertEquals;

@DisplayName("issue I5Q1TW：读取自定义时间格式单元格")
public class IssueI5Q1TWTest {

    @Test
    @DisplayName("读取自定义时间格式单元格并校验时间值")
    public void readTest() {
        final ExcelReader reader = ExcelUtil.getReader("I5Q1TW.xlsx");

        // 自定义时间格式1
        assertEquals("18:56", reader.readCellValue(0, 0).toString());

        // 自定义时间格式2
        assertEquals("18:56", reader.readCellValue(1, 0).toString());
    }
}
