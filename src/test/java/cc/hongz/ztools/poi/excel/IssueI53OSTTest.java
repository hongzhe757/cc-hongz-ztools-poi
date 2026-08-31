package cc.hongz.ztools.poi.excel;

import cc.hongz.ztools.core.lang.Console;
import org.junit.jupiter.api.Disabled;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Sax方式读取合并单元格，只有第一个单元格有值，其余为null。
 */
@DisplayName("Sax 方式读取合并单元格")
public class IssueI53OSTTest {

    @Test
    @Disabled
    @DisplayName("Sax 方式读取合并单元格并解析表头与数据")
    public void readTest() {
        Map<String, Object> result = new HashMap<>();
        List<Object> header = new ArrayList<>();

        ExcelUtil.readBySax("d:/test/sax_merge.xlsx", -1, (sheetIndex, rowIndex, rowCells) -> {
            if (rowIndex == 0) {
                header.addAll(rowCells);
                return;
            }
            for (int i = 0; i < rowCells.size(); i++) {
                result.put((String) header.get(i), rowCells.get(i));
            }
            Console.log(result);
            result.clear();
        });
    }
}
