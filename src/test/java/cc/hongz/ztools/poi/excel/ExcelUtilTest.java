package cc.hongz.ztools.poi.excel;

import cc.hongz.ztools.poi.excel.cell.CellLocation;
import cc.hongz.ztools.poi.excel.sax.handler.RowHandler;
import static org.junit.jupiter.api.Assertions.assertEquals;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import java.util.List;
import java.util.Map;
import java.util.concurrent.atomic.AtomicInteger;

@DisplayName("Excel工具类测试")
public class ExcelUtilTest {

    @DisplayName("列索引转列名测试")
    @Test
    public void indexToColNameTest() {
        assertEquals("A", ExcelUtil.indexToColName(0));
        assertEquals("B", ExcelUtil.indexToColName(1));
        assertEquals("C", ExcelUtil.indexToColName(2));

        assertEquals("AA", ExcelUtil.indexToColName(26));
        assertEquals("AB", ExcelUtil.indexToColName(27));
        assertEquals("AC", ExcelUtil.indexToColName(28));

        assertEquals("AAA", ExcelUtil.indexToColName(702));
        assertEquals("AAB", ExcelUtil.indexToColName(703));
        assertEquals("AAC", ExcelUtil.indexToColName(704));
    }

    @DisplayName("列名转列索引测试")
    @Test
    public void colNameToIndexTest() {
        assertEquals(704, ExcelUtil.colNameToIndex("AAC"));
        assertEquals(703, ExcelUtil.colNameToIndex("AAB"));
        assertEquals(702, ExcelUtil.colNameToIndex("AAA"));

        assertEquals(28, ExcelUtil.colNameToIndex("AC"));
        assertEquals(27, ExcelUtil.colNameToIndex("AB"));
        assertEquals(26, ExcelUtil.colNameToIndex("AA"));

        assertEquals(2, ExcelUtil.colNameToIndex("C"));
        assertEquals(1, ExcelUtil.colNameToIndex("B"));
        assertEquals(0, ExcelUtil.colNameToIndex("A"));
    }

    @DisplayName("单元格地址转坐标测试")
    @Test
    public void toLocationTest() {
        final CellLocation a11 = ExcelUtil.toLocation("A11");
        assertEquals(0, a11.getX());
        assertEquals(10, a11.getY());
    }

    @DisplayName("读取并写出Excel测试")
    @Test
    public void readAndWriteTest() {
        final ExcelReader reader = ExcelUtil.getReader("aaa.xlsx");
        final ExcelWriter writer = reader.getWriter();
        writer.writeCellValue(1, 2, "设置值");
        writer.close();
    }

    @DisplayName("按工作簿路径与Sheet名读取测试")
    @Test
    public void getReaderByBookFilePathAndSheetNameTest() {
        final ExcelReader reader = ExcelUtil.getReader("aaa.xlsx", "12");
        final List<Map<String, Object>> list = reader.readAll();
        reader.close();
        assertEquals(1L, list.get(1).get("鞋码"));
    }

    @DisplayName("SAX读取完成回调测试")
    @Test
    public void doAfterAllAnalysedTest() {
        final String path = "readBySax.xls";
        final AtomicInteger doAfterAllAnalysedTime = new AtomicInteger(0);
        ExcelUtil.readBySax(path, -1, new RowHandler() {
            @Override
            public void handle(final int sheetIndex, final long rowIndex, final List<Object> rowCells) {
                //Console.log("sheetIndex = {}; rowIndex = {},rowCells = {}",sheetIndex,rowIndex,rowCells);
            }

            @Override
            public void doAfterAllAnalysed() {
                doAfterAllAnalysedTime.addAndGet(1);
            }
        });
        //总共2个sheet页，读取所有sheet时，一共执行doAfterAllAnalysed2次。
        assertEquals(2, doAfterAllAnalysedTime.intValue());
    }

}
