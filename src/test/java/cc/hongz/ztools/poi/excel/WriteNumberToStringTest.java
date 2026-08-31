package cc.hongz.ztools.poi.excel;

import cc.hongz.ztools.core.collection.ListUtil;
import cc.hongz.ztools.poi.excel.cell.FormulaCellValue;
import org.junit.jupiter.api.Disabled;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

@DisplayName("数字以文本形式写出测试")
public class WriteNumberToStringTest {
    @Test
    @Disabled
    @DisplayName("写出数字类型数据并保持文本格式")
    public void writeNumberTest() {
        final ExcelWriter writer = ExcelUtil.getBigWriter("d:/test/dataWithNumber.xlsx");

        writer.writeRow(ListUtil.of("姓名", "编号"));
        writer.writeRow(ListUtil.of("张三", new FormulaCellValue("010001")));
        writer.writeRow(ListUtil.of("李四", new FormulaCellValue("120001")));
        writer.writeRow(ListUtil.of("王五", 123456));

        writer.close();
    }
}
