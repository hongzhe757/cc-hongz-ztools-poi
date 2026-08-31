package cc.hongz.ztools.poi.excel;

import cc.hongz.ztools.core.lang.Console;
import org.junit.jupiter.api.Disabled;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import java.util.List;

@DisplayName("issue#3698：Excel 读取")
public class Issue3698Test {
    @Test
    @Disabled
    @DisplayName("读取 Excel 并打印每行第二列值的类型")
    void readTest() {
        ExcelReader reader = ExcelUtil.getReader("d:/test/default.xlsx", 0);
        List<List<Object>> list = reader.read(0, Integer.MAX_VALUE, false);
        for (List<Object> row : list) {
            Console.log(row.get(1).getClass());
        }
        reader.close();
    }
}
