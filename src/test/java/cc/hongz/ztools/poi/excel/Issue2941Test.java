package cc.hongz.ztools.poi.excel;

import cc.hongz.ztools.core.lang.Console;
import org.junit.jupiter.api.Disabled;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import java.util.List;
import java.util.Map;

@DisplayName("issue#2941：Excel 读取日期")
public class Issue2941Test {
    @Test
    @Disabled
    @DisplayName("读取包含日期的 Excel 并打印结果")
    public void excelReadDateTest() {
        final ExcelReader reader = ExcelUtil.getReader("d:/test/1677649503673.xlsx");
        final List<Map<String, Object>> maps = reader.readAll();
        for (Map<String, Object> map : maps) {
            Console.log(map);
        }
    }
}
