package cc.hongz.ztools.poi.excel.sax;

import cc.hongz.ztools.core.lang.Console;
import cc.hongz.ztools.poi.excel.ExcelUtil;
import cc.hongz.ztools.poi.excel.cell.FormulaCellValue;
import org.junit.jupiter.api.Disabled;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

@DisplayName("issue#4195：Sax 方式读取公式单元格")
public class Issue4195Test {
    @Test
    @Disabled
    @DisplayName("Sax 方式读取公式单元格并输出公式与结果")
    void saxReadFormulaTest() {
        // 测试公式读取
        ExcelUtil.readBySax("formula_test.xlsx", -1, (sheetIndex, rowIndex, rowCells) -> {
            final Object value = rowCells.get(2);
            if (value instanceof FormulaCellValue) {
                final FormulaCellValue result = ((FormulaCellValue) value);
                Console.log("公式 {} 结果: {}", result.getValue(), result.getResult());
            } else {
                Console.log("非公式: {}", value.getClass());
            }
        });
    }
}
