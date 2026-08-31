package cc.hongz.ztools.poi.excel;

import cc.hongz.ztools.core.collection.ListUtil;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.junit.jupiter.api.Disabled;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

@DisplayName("issue I64P2K：写出时设置列样式")
public class IssueI64P2KTest {

    @Test
    @Disabled
    @DisplayName("写出数据并设置列样式")
    public void writeWithColumnStyleTest() {
        // 设置默认列样式无效，暂时无解。
        final ExcelWriter writer = ExcelUtil.getWriter("d:/test/columnColorTest.xlsx");
        writer.disableDefaultStyle();

        final Font font = writer.createFont();
        font.setColor(Font.COLOR_RED);
        final CellStyle style = writer.createColumnStyle(0);
        style.setFont(font);
        style.setFillBackgroundColor(IndexedColors.YELLOW.getIndex());

        writer.writeRow(ListUtil.toList("aaa"));
        writer.writeRow(ListUtil.toList("aaa"));

        writer.close();
    }
}
