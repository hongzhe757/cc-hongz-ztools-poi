package cc.hongz.ztools.poi.excel.sax;

import cc.hongz.ztools.core.collection.CollUtil;
import cc.hongz.ztools.core.convert.Convert;
import cc.hongz.ztools.core.io.FileUtil;
import cc.hongz.ztools.core.io.IoUtil;
import cc.hongz.ztools.core.lang.Console;
import cc.hongz.ztools.core.util.StrUtil;
import cc.hongz.ztools.poi.excel.ExcelUtil;
import cc.hongz.ztools.poi.excel.cell.FormulaCellValue;
import cc.hongz.ztools.poi.excel.sax.handler.RowHandler;
import cc.hongz.ztools.poi.exceptions.POIException;
import org.apache.poi.ss.usermodel.CellStyle;
import org.junit.jupiter.api.Disabled;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * Excel sax方式读取。
 *
 * @author looly
 */
@DisplayName("Excel SAX 方式读取测试")
public class ExcelSaxReadTest {

    @Test
    @DisplayName("读取 xlsx 文件（按 sheet 索引）")
    public void excel07Test() {
        // 工具化快速读取
        ExcelUtil.readBySax("aaa.xlsx", 0, createRowHandler());
    }

    @Test
    @DisplayName("通过异常提前终止读取（xlsx）")
    void readEndByExceptionTest() {
        ExcelUtil.readBySax("aaa.xlsx", 0, (sheetIndex, rowIndex, rowList) -> {
            if (rowIndex == 1) {
                throw new StopReadException();
            }
        });
    }

    @Test
    @DisplayName("通过异常提前终止读取（xls）")
    void readEndByException03Test() {
        ExcelUtil.readBySax("aaa.xls", 0, (sheetIndex, rowIndex, rowList) -> {
            if (rowIndex == 1) {
                throw new StopReadException();
            }
        });
    }

    @Test
    @DisplayName("按 sheet 名称读取 xlsx 文件")
    public void excel07ByNameTest() {
        // 工具化快速读取
        // sheet名称是区分大小写的
        ExcelUtil.readBySax("aaa.xlsx", "Sheet1", createRowHandler());
        // 纯数字名称也支持
        ExcelUtil.readBySax("aaa.xlsx", "12", createRowHandler());
        // 前缀支持
        ExcelUtil.readBySax("aaa.xlsx", "sheetName:12", createRowHandler());
    }

    @Test
    @DisplayName("从输入流读取 xlsx 文件")
    public void excel07FromStreamTest() {
        // issue#1225 非markSupport的流读取会错误
        ExcelUtil.readBySax(IoUtil.toStream(FileUtil.file("aaa.xlsx")), 0, createRowHandler());
    }

    @Test
    @DisplayName("读取 xls 文件（按 sheet 索引）")
    public void excel03Test() {
        final Excel03SaxReader reader = new Excel03SaxReader(createRowHandler());
        reader.read("aaa.xls", 1);

        // Console.log("Sheet index: [{}], Sheet name: [{}]", reader.getSheetIndex(), reader.getSheetName());
        ExcelUtil.readBySax("aaa.xls", 1, createRowHandler());
    }

    @Test
    @DisplayName("按 sheet 名称读取 xls 文件")
    public void excel03ByNameTest() {
        final Excel03SaxReader reader = new Excel03SaxReader(createRowHandler());
        reader.read("aaa.xls", "校园入学");
        reader.read("aaa.xls", "sheetName:校园入学");
    }

    @Test
    @DisplayName("sheet 名称不存在时读取抛出异常")
    public void excel03ByNameErrorTest() {
        assertThrows(POIException.class, () -> {
            // sheet名称不存在则报错
            final Excel03SaxReader reader = new Excel03SaxReader(createRowHandler());
            reader.read("aaa.xls", "校园入学1");
        });
    }

    @Test
    @Disabled
    @DisplayName("读取空白行文件")
    public void readBlankLineTest() {
        ExcelUtil.readBySax("e:/ExcelBlankLine.xlsx", 0, (sheetIndex, rowIndex, rowList) -> {
            if (StrUtil.isAllEmpty(Convert.toStrArray(rowList))) {
                return;
            }
            Console.log(rowList);
        });
    }

    @Test
    @DisplayName("按 sheet 名称（纯数字）读取")
    public void readBySaxTest() {
        ExcelUtil.readBySax("blankAndDateTest.xlsx", "0", createRowHandler());
    }

    @Test
    @DisplayName("按 sheet 索引读取")
    public void readBySaxByRidTest() {
        ExcelUtil.readBySax("blankAndDateTest.xlsx", 0, createRowHandler());
    }

    @Test
    @DisplayName("按 sheet 名称读取")
    public void readBySaxByNameTest() {
        ExcelUtil.readBySax("blankAndDateTest.xlsx", "Sheet1", createRowHandler());
    }

    @Test
    @Disabled
    @DisplayName("读取指定路径的 xlsx 文件")
    public void readBySaxTest2() {
        ExcelUtil.readBySax("d:/test/456789.xlsx", "0", (sheetIndex, rowIndex, rowList) -> Console.log(rowList));
    }

    private RowHandler createRowHandler() {
        return (sheetIndex, rowIndex, rowlist) -> {
//            Console.log("[{}] [{}] {}", sheetIndex, rowIndex, rowlist);
            if (5 != rowIndex && 6 != rowIndex) {
                // 测试样例中除第五行、第六行都为非空行
                assertTrue(CollUtil.isNotEmpty(rowlist));
            }
        };
    }

    @Test
    @Disabled
    @DisplayName("逐单元格读取 xlsx 文件（handleCell）")
    public void handle07CellTest() {
        ExcelUtil.readBySax("d:/test/test.xlsx", -1, new RowHandler() {

                    @Override
                    public void handleCell(final int sheetIndex, final long rowIndex, final int cellIndex,
                            final Object value, final CellStyle xssfCellStyle) {
                        Console.log("{} {} {}", rowIndex, cellIndex, value);
                    }

                    @Override
                    public void handle(final int sheetIndex, final long rowIndex, final List<Object> rowCells) {

                    }
                }
        );
    }

    @Test
    @DisplayName("逐单元格读取 xls 文件（handleCell）")
    public void handle03CellTest() {
        ExcelUtil.readBySax("test.xls", -1, new RowHandler() {

                    @Override
                    public void handleCell(final int sheetIndex, final long rowIndex, final int cellIndex,
                            final Object value, final CellStyle xssfCellStyle) {
                        //Console.log("{} {} {}", rowIndex, cellIndex, value);
                    }

                    @Override
                    public void handle(final int sheetIndex, final long rowIndex, final List<Object> rowCells) {
                    }
                }
        );
    }

    @Test
    @DisplayName("读取 xls 公式单元格")
    public void formulaRead03Test() {
        final List<Object> rows = new ArrayList<>();
        ExcelUtil.readBySax("data_for_sax_test.xls", -1, (i, i1, list) -> {
            if (list.size() > 1) {
                rows.add(list.get(1));
            } else {
                rows.add("");
            }
        });
        assertEquals(50L, rows.get(3));
    }

    @Test
    @DisplayName("读取 xlsx 公式单元格")
    public void formulaRead07Test() {
        final List<Object> rows = new ArrayList<>();
        ExcelUtil.readBySax("data_for_sax_test.xlsx", 0, (i, i1, list) -> rows.add(list.get(1)));

        final FormulaCellValue value = (FormulaCellValue) rows.get(3);
        assertEquals(50L, value.getResult());
    }

    @Test
    @DisplayName("读取 xls 日期单元格")
    public void dateReadXlsTest() {
        final List<String> rows = new ArrayList<>();
        ExcelUtil.readBySax("data_for_sax_test.xls", 0,
                (i, i1, list) -> rows.add(StrUtil.toString(list.get(0)))
        );

        assertEquals("2020-10-09 00:00:00", rows.get(1));
        // 非日期格式不做转换
        assertEquals("112233", rows.get(2));
        assertEquals("1000.0", rows.get(3));
        assertEquals("2012-12-21 00:00:00", rows.get(4));
    }

    @Test
    @DisplayName("读取 xlsx 日期单元格")
    public void dateReadXlsxTest() {
        final List<String> rows = new ArrayList<>();
        ExcelUtil.readBySax("data_for_sax_test.xlsx", 0,
                (i, i1, list) -> rows.add(StrUtil.toString(list.get(0)))
        );

        assertEquals("2020-10-09 00:00:00", rows.get(1));
        // 非日期格式不做转换
        assertEquals("112233", rows.get(2));
        assertEquals("1000.0", rows.get(3));
        assertEquals("2012-12-21 00:00:00", rows.get(4));
    }

    @Test
    @Disabled
    @DisplayName("读取自定义日期格式文件")
    public void dateReadXlsxTest2() {
        ExcelUtil.readBySax("d:/test/custom_date_format2.xlsx", 0,
                (i, i1, list) -> Console.log(list)
        );
    }

    @Test
    @Disabled
    @DisplayName("读取空白文件")
    public void readBlankTest() {
        final File file = new File("D:/test/b.xlsx");

        ExcelUtil.readBySax(file, 0, (sheetIndex, rowIndex, rowList) -> rowList.forEach(Console::log));

        ExcelUtil.getReader(file).read().forEach(Console::log);
    }

    @Test
    @Disabled
    @DisplayName("读取 xlsm 文件")
    public void readXlsmTest() {
        ExcelUtil.readBySax("d:/test/WhiteListTemplate.xlsm", -1,
                (sheetIndex, rowIndex, rowlist) -> Console.log("[{}] [{}] {}", sheetIndex, rowIndex, rowlist));
    }
}
