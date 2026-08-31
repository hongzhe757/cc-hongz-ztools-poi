package cc.hongz.ztools.poi.excel;

import cc.hongz.ztools.core.collection.CollUtil;
import cc.hongz.ztools.core.date.DateUtil;
import cc.hongz.ztools.poi.excel.cell.CellSetter;
import cc.hongz.ztools.poi.excel.style.StyleUtil;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.junit.jupiter.api.Disabled;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import java.util.List;

/**
 * 设置列样式测试。
 */
@DisplayName("设置列样式写出测试")
public class WriteStyleTest {
    @Test
    @Disabled
    @DisplayName("写出数据并设置自定义字体与列样式")
    public void writeTest() {
        List<?> row1 = CollUtil.newArrayList("aaaaa", "bb", "cc", "dd", DateUtil.date(), 3.22676575765);
        List<?> row2 = CollUtil.newArrayList("aa1", "bb1", "cc1", "dd1", DateUtil.date(), 250.7676);
        List<?> row3 = CollUtil.newArrayList("aa2", "bb2", "cc2", "dd2", DateUtil.date(), 0.111);
        List<?> row4 = CollUtil.newArrayList("aa3", "bb3", "cc3", "dd3", DateUtil.date(), 35);
        List<?> row5 = CollUtil.newArrayList("aa4", "bb4", "cc4", "dd4", DateUtil.date(), 28.00);
        List<List<?>> rows = CollUtil.newArrayList(row1, row2, row3, row4, row5);

        ExcelWriter writer = ExcelUtil.getWriter("d:/test/writeTest.xlsx");
        writer.setStyleSet(null);

        final CellStyle cellStyle = StyleUtil.createDefaultCellStyle(writer.getWorkbook());
        cellStyle.setFont(StyleUtil.createFont(writer.getWorkbook(), Font.COLOR_RED, (short) 12, "宋体"));

        writer.write(rows);

        writer.writeCellValue(0, 0, (CellSetter) cell -> {
            cell.setCellStyle(cellStyle);
            cell.setCellValue("1234");
        });

        for (int i = 0; i < writer.getRowCount(); i++) {
            writer.setStyle(cellStyle, 2, i);
        }

        writer.close();
    }
}
