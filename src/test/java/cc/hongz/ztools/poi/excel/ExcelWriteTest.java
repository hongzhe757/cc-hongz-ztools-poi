package cc.hongz.ztools.poi.excel;

import cc.hongz.ztools.core.collection.CollUtil;
import cc.hongz.ztools.core.collection.ListUtil;
import cc.hongz.ztools.core.comparator.IndexedComparator;
import cc.hongz.ztools.core.date.DateUtil;
import cc.hongz.ztools.core.io.FileUtil;
import cc.hongz.ztools.core.lang.Console;
import cc.hongz.ztools.core.map.MapUtil;
import cc.hongz.ztools.core.util.CharsetUtil;
import cc.hongz.ztools.core.util.IdUtil;
import cc.hongz.ztools.core.util.ObjectUtil;
import cc.hongz.ztools.poi.excel.cell.setters.EscapeStrCellSetter;
import cc.hongz.ztools.poi.excel.style.StyleUtil;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import static org.junit.jupiter.api.Assertions.assertEquals;
import org.junit.jupiter.api.Disabled;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;
import java.util.TreeMap;

/**
 * 写出Excel单元测试。
 *
 * @author looly
 */
@DisplayName("写出 Excel 单元测试")
public class ExcelWriteTest {

    /** 测试数据1。 */
    public static Map<String, Object> MAP_DATA_1 = new LinkedHashMap<>();
    /** 测试数据2。 */
    public static Map<String, Object> MAP_DATA_2 = new LinkedHashMap<>();

    static {
        MAP_DATA_1.put("姓名", "张三");
        MAP_DATA_1.put("年龄", 23);
        MAP_DATA_1.put("成绩", 88.32);
        MAP_DATA_1.put("是否合格", true);
        MAP_DATA_1.put("考试日期", DateUtil.date());

        MAP_DATA_2.put("姓名", "张三");
        MAP_DATA_2.put("年龄", 23);
        MAP_DATA_2.put("成绩", 88.32);
        MAP_DATA_2.put("是否合格", true);
        MAP_DATA_2.put("考试日期", DateUtil.date());
    }

    @DisplayName("测试写出多行数据（不显式刷新）")
    @Test
    public void writeNoFlushTest() {
        List<?> row1 = CollUtil.newArrayList("aaaaa", "bb", "cc", "dd", DateUtil.date(), 3.22676575765);
        List<?> row2 = CollUtil.newArrayList("aa1", "bb1", "cc1", "dd1", DateUtil.date(), 250.7676);
        List<?> row3 = CollUtil.newArrayList("aa2", "bb2", "cc2", "dd2", DateUtil.date(), 0.111);
        List<?> row4 = CollUtil.newArrayList("aa3", "bb3", "cc3", "dd3", DateUtil.date(), 35);
        List<?> row5 = CollUtil.newArrayList("aa4", "bb4", "cc4", "dd4", DateUtil.date(), 28.00);
        List<List<?>> rows = CollUtil.newArrayList(row1, row2, row3, row4, row5);

        final ExcelWriter writer = ExcelUtil.getWriter();
        writer.write(rows);
        writer.close();
    }

    @DisplayName("测试设置行和列的单元格样式")
    @Test
    @Disabled
    public void testRowOrColumnCellStyle() {
        List<?> row1 = CollUtil.newArrayList("aaaaa", "bb", "cc", "dd", DateUtil.date(), 3.22676575765);
        List<?> row2 = CollUtil.newArrayList("aa1", "bb1", "cc1", "dd1", DateUtil.date(), 250.7676);
        List<?> row3 = CollUtil.newArrayList("aa2", "bb2", "cc2", "dd2", DateUtil.date(), 0.111);
        List<?> row4 = CollUtil.newArrayList("aa3", "bb3", "cc3", "dd3", DateUtil.date(), 35);
        List<?> row5 = CollUtil.newArrayList("aa4", "bb4", "cc4", "dd4", DateUtil.date(), 28.00);

        List<List<?>> rows = CollUtil.newArrayList(row1, row2, row3, row4, row5);
        BigExcelWriter overtimeWriter = ExcelUtil.getBigWriter("d:/test/style_line.xlsx");

        overtimeWriter.write(rows, true);

        CellStyle cellStyle = overtimeWriter.getWorkbook().createCellStyle();
        StyleUtil.setBorder(cellStyle, BorderStyle.THIN, IndexedColors.BLACK);
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        cellStyle.setFillForegroundColor((short) 13);
        cellStyle.setDataFormat((short) 22); //时间格式
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        //原设置行、列样式的方法
//        overtimeWriter.setRowStyle(2,cellStyle);
//        overtimeWriter.setColumnStyle(1,cellStyle);

        //现增加的设置行、列样式的方法
        //给第三行加背景色
        overtimeWriter.setRowStyleIfHasData(2, cellStyle);
        //给第二列加背景色 从第一行开始加（用于控制有表头时）
        overtimeWriter.setColumnStyleIfHasData(1, 0, cellStyle);

        CellStyle cellStyle1 = overtimeWriter.getWorkbook().createCellStyle();
        StyleUtil.setBorder(cellStyle1, BorderStyle.THIN, IndexedColors.BLACK);
        cellStyle1.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        cellStyle1.setFillForegroundColor((short) 13);
        cellStyle1.setDataFormat((short) 2); //小数保留两位
        cellStyle1.setAlignment(HorizontalAlignment.CENTER);
        cellStyle1.setVerticalAlignment(VerticalAlignment.CENTER);
        overtimeWriter.setStyle(cellStyle1, 5, 2); //由于第6列是数字 上面应用了日期格式会错乱，这里单独设置下第六列的格式

        overtimeWriter.close();
    }

    @DisplayName("测试写出单行数据到文件")
    @Test
    @Disabled
    public void writeTest2() {
        List<String> row = CollUtil.newArrayList("姓名", "加班日期", "下班时间", "加班时长", "餐补", "车补次数", "车补", "总计");
        ExcelWriter overtimeWriter = ExcelUtil.getWriter("e:/excel/single_line.xlsx");
        overtimeWriter.writeRow(row);
        overtimeWriter.close();
    }

    @DisplayName("测试写出到多个 Sheet 页")
    @Test
    @Disabled
    public void writeWithSheetTest() {
        ExcelWriter writer = ExcelUtil.getWriterWithSheet("表格1");

        // 写出第一张表
        List<String> row = CollUtil.newArrayList("姓名", "加班日期", "下班时间", "加班时长", "餐补", "车补次数", "车补", "总计");
        writer.writeRow(row);

        // 写出第二张表
        writer.setSheet("表格2");
        List<String> row2 = CollUtil.newArrayList("姓名2", "加班日期2", "下班时间2", "加班时长2", "餐补2", "车补次数2", "车补2", "总计2");
        writer.writeRow(row2);

        // 生成文件或导出Excel
        writer.flush(FileUtil.file("f:/test/writeWithSheetTest.xlsx"));

        writer.close();
    }

    @DisplayName("测试写出超大列表数据（合并标题、冻结行）")
    @Test
    @Disabled
    public void writeTest() {
        List<?> row1 = CollUtil.newArrayList("aaaaa", "bb", "cc", "dd", DateUtil.date(), 3.22676575765);
        List<?> row2 = CollUtil.newArrayList("aa1", "bb1", "cc1", "dd1", DateUtil.date(), 250.7676);
        List<?> row3 = CollUtil.newArrayList("aa2", "bb2", "cc2", "dd2", DateUtil.date(), 0.111);
        List<?> row4 = CollUtil.newArrayList("aa3", "bb3", "cc3", "dd3", DateUtil.date(), 35);
        List<?> row5 = CollUtil.newArrayList("aa4", "bb4", "cc4", "dd4", DateUtil.date(), 28.00);

        List<List<?>> rows = CollUtil.newArrayList(row1, row2, row3, row4, row5);
        for (int i = 0; i < 400; i++) {
            // 超大列表写出测试
            rows.add(ObjectUtil.clone(row1));
        }

        String filePath = "d:/test/writeTest.xlsx";
        FileUtil.del(filePath);
        // 通过工具类创建writer
        ExcelWriter writer = ExcelUtil.getWriter(filePath);
        // 通过构造方法创建writer
        // ExcelWriter writer = new ExcelWriter("d:/writeTest.xls");

        // 跳过当前行，即第一行，非必须，在此演示用
        writer.passCurrentRow();
        // 合并单元格后的标题行，使用默认标题样式
        writer.merge(row1.size() - 1, "测试标题");
        // 一次性写出内容，使用默认样式
        writer.write(rows);
        writer.autoSizeColumn(0, true);
        //冻结前两行
        writer.setFreezePane(0, 2);
        // 关闭writer，释放内存
        writer.close();
    }

    @DisplayName("测试合并单元格")
    @Test
    @Disabled
    public void mergeTest() {
        List<?> row1 = CollUtil.newArrayList("aa", "bb", "cc", "dd", DateUtil.date(), 3.22676575765);
        List<?> row2 = CollUtil.newArrayList("aa1", "bb1", "cc1", "dd1", DateUtil.date(), 250.7676);
        List<?> row3 = CollUtil.newArrayList("aa2", "bb2", "cc2", "dd2", DateUtil.date(), 0.111);
        List<?> row4 = CollUtil.newArrayList("aa3", "bb3", "cc3", "dd3", DateUtil.date(), 35);
        List<?> row5 = CollUtil.newArrayList("aa4", "bb4", "cc4", "dd4", DateUtil.date(), 28.00);

        List<List<?>> rows = CollUtil.newArrayList(row1, row2, row3, row4, row5);

        // 通过工具类创建writer
        ExcelWriter writer = ExcelUtil.getWriter("d:/test/mergeTest.xlsx");
        CellStyle style = writer.getStyleSet().getHeadCellStyle();
        StyleUtil.setColor(style, IndexedColors.RED, FillPatternType.SOLID_FOREGROUND);

        // 跳过当前行，即第一行，非必须，在此演示用
        writer.passCurrentRow();
        // 合并单元格后的标题行，使用默认标题样式
        writer.merge(row1.size() - 1, "测试标题");
        // 一次性写出内容，使用默认样式
        writer.write(rows);

        // 合并单元格后的标题行，使用默认标题样式
        writer.merge(7, 10, 4, 10, "测试Merge", false);

        // 关闭writer，释放内存
        writer.close();
    }

    @DisplayName("测试合并标题并写出 Map 数据")
    @Test
    @Disabled
    public void mergeTest2() {

        ArrayList<Map<String, Object>> rows = CollUtil.newArrayList(MAP_DATA_1, MAP_DATA_2);

        // 通过工具类创建writer
        ExcelWriter writer = ExcelUtil.getWriter("d:/test/writeMapTest.xlsx");
        // 合并单元格后的标题行，使用默认标题样式
        writer.merge(MAP_DATA_1.size() - 1, "一班成绩单");

        // 一次性写出内容，使用默认样式，强制输出标题
        writer.write(rows, true);
        // 关闭writer，释放内存
        writer.close();
    }

    @DisplayName("测试写出 Map 数据并设置自定义字体")
    @Test
    @Disabled
    public void writeMapTest() {

        ArrayList<Map<String, Object>> rows = CollUtil.newArrayList(MAP_DATA_1, MAP_DATA_2);

        // 通过工具类创建writer
        ExcelWriter writer = ExcelUtil.getWriter("e:/excel/writeMapTest.xlsx");

        // 设置内容字体
        Font font = writer.createFont();
        font.setBold(true);
        font.setColor(Font.COLOR_RED);
        font.setItalic(true);
        writer.getStyleSet().setFont(font, true);

        // 合并单元格后的标题行，使用默认标题样式
        writer.merge(MAP_DATA_1.size() - 1, "一班成绩单");
        // 一次性写出内容，使用默认样式
        writer.write(rows, true);
        // 关闭writer，释放内存
        writer.close();
    }

    @DisplayName("测试写出单行 Map 数据")
    @Test
    @Disabled
    public void writeMapTest2() {

        // 通过工具类创建writer
        ExcelWriter writer = ExcelUtil.getWriter("d:/writeMapTest2.xlsx");

        // 一次性写出内容，使用默认样式
        writer.writeRow(MAP_DATA_1, true);
        // 关闭writer，释放内存
        writer.close();
    }

    @DisplayName("测试写出 Map 数据并设置单元格样式")
    @Test
    @Disabled
    public void writeMapWithStyleTest() {

        // 通过工具类创建writer
        String path = "f:/test/writeMapWithStyleTest.xlsx";
        FileUtil.del(path);
        ExcelWriter writer = ExcelUtil.getWriter(path);
        writer.setStyleSet(null);

        // 一次性写出内容，使用默认样式
        writer.writeRow(MAP_DATA_1, true);

        // 设置某个单元格样式
        CellStyle orCreateRowStyle = writer.getOrCreateCellStyle(0, 1);
        StyleUtil.setColor(orCreateRowStyle, IndexedColors.RED.getIndex(), FillPatternType.SOLID_FOREGROUND);

        // 关闭writer，释放内存
        writer.close();
    }

    @DisplayName("测试写出 Map 数据并配置表头别名")
    @Test
    @Disabled
    public void writeMapAliasTest() {
        Map<Object, Object> row1 = new LinkedHashMap<>();
        row1.put("name", "张三");
        row1.put("age", 22);
        row1.put("isPass", true);
        row1.put("score", 66.30);
        row1.put("examDate", DateUtil.date());
        Map<Object, Object> row2 = new LinkedHashMap<>();
        row2.put("name", "李四");
        row2.put("age", 233);
        row2.put("isPass", false);
        row2.put("score", 32.30);
        row2.put("examDate", DateUtil.date());

        List<Map<Object, Object>> rows = CollUtil.newArrayList(row1, row2);
        // 通过工具类创建writer
        String file = "d:/test/writeMapAlias.xlsx";
        FileUtil.del(file);
        ExcelWriter writer = ExcelUtil.getWriter(file);
        // 自定义标题
        writer.addHeaderAlias("name", "姓名");
        writer.addHeaderAlias("age", "年龄");
        writer.addHeaderAlias("score", "分数");
        writer.addHeaderAlias("isPass", "是否合格");
        writer.addHeaderAlias("examDate", "考试日期");
        // 合并单元格后的标题行，使用默认标题样式
        writer.merge(4, "一班成绩单");
        // 一次性写出内容，使用默认样式
        writer.write(rows, true);
        // 关闭writer，释放内存
        writer.close();
    }

    @DisplayName("测试仅输出已配置别名的 Map 字段")
    @Test
    @Disabled
    public void writeMapOnlyAliasTest() {
        Map<Object, Object> row1 = new LinkedHashMap<>();
        row1.put("name", "张三");
        row1.put("age", 22);
        row1.put("isPass", true);
        row1.put("score", 66.30);
        row1.put("examDate", DateUtil.date());
        Map<Object, Object> row2 = new LinkedHashMap<>();
        row2.put("name", "李四");
        row2.put("age", 233);
        row2.put("isPass", false);
        row2.put("score", 32.30);
        row2.put("examDate", DateUtil.date());

        List<Map<Object, Object>> rows = CollUtil.newArrayList(row1, row2);
        // 通过工具类创建writer
        String file = "f:/test/test_alias.xlsx";
        FileUtil.del(file);
        ExcelWriter writer = ExcelUtil.getWriter(file);
        writer.setOnlyAlias(true);
        // 自定义标题
        writer.addHeaderAlias("name", "姓名");
        writer.addHeaderAlias("age", "年龄");
        // 合并单元格后的标题行，使用默认标题样式
        writer.merge(4, "一班成绩单");
        // 一次性写出内容，使用默认样式
        writer.write(rows, true);
        // 关闭writer，释放内存
        writer.close();
    }

    @DisplayName("测试写出 Map 数据到 xls 文件并使用别名")
    @Test
    @Disabled
    public void writeMapOnlyAliasTest2() {
        Map<Object, Object> row1 = new LinkedHashMap<>();
        row1.put("name", "张三");
        row1.put("age", 22);
        row1.put("isPass", true);
        row1.put("score", 66.30);
        row1.put("examDate", DateUtil.date());
        Map<Object, Object> row2 = new LinkedHashMap<>();
        row2.put("name", "李四");
        row2.put("age", 233);
        row2.put("isPass", false);
        row2.put("score", 32.30);
        row2.put("examDate", DateUtil.date());

        List<Map<Object, Object>> rows = CollUtil.newArrayList(row1, row2);
        // 通过工具类创建writer
        String file = "d:/test/test_alias.xls";
        ExcelWriter writer = ExcelUtil.getWriter(file, "test1");
//        writer.setOnlyAlias(true);
        // 自定义标题
        writer.addHeaderAlias("name", "姓名");
        writer.addHeaderAlias("age", "年龄");
        // 一次性写出内容，使用默认样式
        writer.write(rows, true);
        // 关闭writer，释放内存
        writer.close();
    }

    @DisplayName("测试仅输出别名且部分字段缺失时的写出")
    @Test
    @Disabled
    public void writeMapOnlyAliasTest3() {
        Map<Object, Object> row1 = new LinkedHashMap<>();
        row1.put("name", "张三");
        row1.put("age", 22);
        row1.put("isPass", true);
        row1.put("score", 66.30);
        row1.put("examDate", DateUtil.date());

        Map<Object, Object> row2 = new LinkedHashMap<>();
        row2.put("name", "李四");
//        row2.put("age", 233);
        row2.put("isPass", false);
        row2.put("score", 32.30);
        row2.put("examDate", DateUtil.date());

        List<Map<Object, Object>> rows = CollUtil.newArrayList(row1, row2);
        // 通过工具类创建writer
        String file = "d:/test/test_alias.xls";
        ExcelWriter writer = ExcelUtil.getWriter(file, "test1");
        writer.setOnlyAlias(true);

        // 自定义标题
        writer.addHeaderAlias("name", "姓名");
        writer.addHeaderAlias("age", "年龄");
        writer.addHeaderAlias("examDate", "考试时间");

        // 一次性写出内容，使用默认样式
        writer.write(rows, true);
        // 关闭writer，释放内存
        writer.close();
    }

    @DisplayName("测试写出 Bean 列表并配置表头别名")
    @Test
    @Disabled
    public void writeBeanTest() {
        TestBean bean1 = new TestBean();
        bean1.setName("张三");
        bean1.setAge(22);
        bean1.setPass(true);
        bean1.setScore(66.30);
        bean1.setExamDate(DateUtil.date());

        TestBean bean2 = new TestBean();
        bean2.setName("李四");
        bean2.setAge(28);
        bean2.setPass(false);
        bean2.setScore(38.50);
        bean2.setExamDate(DateUtil.date());

        List<TestBean> rows = CollUtil.newArrayList(bean1, bean2);
        // 通过工具类创建writer
        String file = "e:/writeBeanTest.xlsx";
        FileUtil.del(file);
        ExcelWriter writer = ExcelUtil.getWriter(file);
        // 自定义标题
        writer.addHeaderAlias("name", "姓名");
        writer.addHeaderAlias("age", "年龄");
        writer.addHeaderAlias("score", "分数");
        writer.addHeaderAlias("isPass", "是否通过");
        writer.addHeaderAlias("examDate", "考试时间");
        // 合并单元格后的标题行，使用默认标题样式
        writer.merge(4, "一班成绩单");
        // 一次性写出内容，使用默认样式
        writer.write(rows, true);
        // 关闭writer，释放内存
        writer.close();
    }

    @DisplayName("测试写出 OrderExcel 实体列表")
    @Test
    @Disabled
    public void writeBeanTest2() {
        OrderExcel order1 = new OrderExcel();
        order1.setId("1");
        order1.setNum("123");
        order1.setBody("body1");

        OrderExcel order2 = new OrderExcel();
        order1.setId("2");
        order1.setNum("456");
        order1.setBody("body2");

        List<OrderExcel> rows = CollUtil.newArrayList(order1, order2);
        // 通过工具类创建writer
        String file = "f:/test/writeBeanTest2.xlsx";
        FileUtil.del(file);
        ExcelWriter writer = ExcelUtil.getWriter(file);
        // 自定义标题
        writer.addHeaderAlias("id", "编号");
        writer.addHeaderAlias("num", "序号");
        writer.addHeaderAlias("body", "内容");
        // 一次性写出内容，使用默认样式
        writer.write(rows, true);
        // 关闭writer，释放内存
        writer.close();
    }

    @DisplayName("测试写出指定单元格的值")
    @Test
    @Disabled
    public void writeCellValueTest() {
        ExcelWriter writer = new ExcelWriter("d:/cellValueTest.xls");
        writer.writeCellValue(3, 5, "aaa");
        writer.writeCellValue(3, 5, "aaa");
        writer.close();
    }

    @DisplayName("测试添加单元格下拉选择框")
    @Test
    @Disabled
    public void addSelectTest() {
        List<String> row = CollUtil.newArrayList("姓名", "加班日期", "下班时间", "加班时长", "餐补", "车补次数", "车补", "总计");
        ExcelWriter overtimeWriter = ExcelUtil.getWriter("d:/test/single_line.xlsx");
        overtimeWriter.writeCellValue(3, 4, "AAAA");
        overtimeWriter.addSelect(3, 4, row.toArray(new String[0]));
        overtimeWriter.close();
    }

    @DisplayName("测试按单元格区域添加下拉选择框")
    @Test
    @Disabled
    public void addSelectTest2() {
        ExcelWriter writer = ExcelUtil.getWriter("d:/test/select.xls");
        writer.writeCellValue(0, 0, "请选择科目");
        int firstRow = 0;
        int lastRow = 0;
        int firstCol = 0;
        int lastCol = 0;
        CellRangeAddressList addressList = new CellRangeAddressList(firstRow, lastRow, firstCol, lastCol);
        writer.addSelect(addressList, "1001", "1002", "1003");

        List<?> rows = new ArrayList<>();
        writer.write(rows, true);

        writer.close();
    }

    @DisplayName("测试写出多个 Sheet 页（含别名和自适应列宽）")
    @Test
    @Disabled
    public void writeMultiSheetTest() {
        List<Map<String, Object>> rows = new LinkedList<>();
        for (int i = 0; i < 10; i++) {
            Map<String, Object> tempList = new TreeMap<>();
            for (int j = 0; j < 10; j++) {
                tempList.put(j + "", IdUtil.randomUUID());
            }
            rows.add(tempList);
        }
        ExcelWriter writer = ExcelUtil.getWriter("D:\\test\\multiSheet.xlsx", "正常数据");
        writer.addHeaderAlias("1", "row1");
        writer.addHeaderAlias("3", "row2");
        writer.setOnlyAlias(true);

        writer.write(rows, true);
        writer.autoSizeColumnAll();

        //表2
        writer.setSheet("当前重复数据");
        writer.clearHeaderAlias();
        writer.addHeaderAlias("3", "行3");
        writer.addHeaderAlias("1", "行1");
        writer.write(rows, true);
        writer.autoSizeColumnAll();

        //表3
        writer.setSheet("历史重复数据");
        writer.write(rows, true);
        writer.autoSizeColumnAll();

        writer.close();
    }

    @DisplayName("测试写出多个 Sheet 页")
    @Test
    @Disabled
    public void writeMultiSheetTest2() {
        List<Map<String, Object>> rows = new LinkedList<>();
        final HashMap<String, Object> map = MapUtil.newHashMap();
        map.put("k1", "v1");
        map.put("k2", "v2");
        map.put("k3", "v3");
        rows.add(map);

        ExcelWriter writer = ExcelUtil.getWriter("D:\\test\\multiSheet2.xlsx", "正常数据");
        writer.write(rows);

        //表2
        writer.setSheet("表2");
        List<Map<String, Object>> rows2 = new LinkedList<>();
        final HashMap<String, Object> map2 = MapUtil.newHashMap();
        map2.put("x1", "v1");
        rows2.add(map2);
        writer.write(rows2);

        writer.close();
    }

    @DisplayName("测试写出多个 Sheet 页并设置表头样式")
    @Test
    @Disabled
    public void writeMultiSheetWithStyleTest() {
        ExcelWriter writer = ExcelUtil.getWriter("D:\\test\\multiSheetWithStyle.xlsx", "表格1");

        // 表1
        List<Map<String, Object>> rows = new LinkedList<>();
        final HashMap<String, Object> map = MapUtil.newHashMap();
        map.put("k1", "v1");
        map.put("k2", "v2");
        map.put("k3", "v3");
        rows.add(map);
        writer.write(rows);

        Font headFont = writer.createFont();
        headFont.setBold(true);
        headFont.setFontHeightInPoints((short) 50);
        headFont.setFontName("Microsoft YaHei");
        writer.getStyleSet().getHeadCellStyle().setFont(headFont);

        //表2
        writer.setSheet("表2");
        List<Map<String, Object>> rows2 = new LinkedList<>();
        final HashMap<String, Object> map2 = MapUtil.newHashMap();
        map2.put("x1", "v1");
        rows2.add(map2);
        writer.write(rows2);

        writer.close();
    }

    @DisplayName("测试写出键数量不一致的 Map 列表")
    @Test
    @Disabled
    public void writeMapsTest() {
        List<Map<String, Object>> rows = new ArrayList<>();

        Map<String, Object> map1 = new HashMap<>();
        map1.put("a", 1);
        map1.put("b", 2);
        map1.put("c", 3);
        map1.put("d", 4);
        map1.put("e", 5);
        Map<String, Object> map2 = new HashMap<>();
        map2.put("c", 3);
        map2.put("d", 4);
        map2.put("e", 5);
        Map<String, Object> map3 = new HashMap<>();
        map3.put("d", 4);
        map3.put("e", 5);

        rows.add(map1);
        rows.add(map2);
        rows.add(map3);

        final ExcelWriter writer = ExcelUtil.getWriter("d:/test/rows.xlsx");
        writer.write(rows);
        writer.close();
    }

    @DisplayName("测试设置单元格日期格式")
    @Test
    @Disabled
    public void formatTest() {
        final ExcelWriter writer = ExcelUtil.getWriter("d:/test/formatTest.xlsx");
        final CellStyle cellStyle = writer.createCellStyle(0, 0);
        cellStyle.setDataFormat(writer.getWorkbook().createDataFormat().getFormat("yyyy-mm-dd"));
        writer.close();
    }

    @DisplayName("测试写出数字并设置两位小数格式")
    @Test
    @Disabled
    public void writeNumberFormatTest() {
        final ExcelWriter writer = ExcelUtil.getWriter("d:/test/formatTest.xlsx");
        writer.disableDefaultStyle();
        writer.writeRow(ListUtil.toList(51.33333333, 90.111111111));
        final CellStyle columnStyle = writer.createCellStyle(0, 0);
        columnStyle.setDataFormat((short) BuiltinFormats.getBuiltinFormat("0.00"));
        writer.close();
    }

    @DisplayName("测试写出复杂表头（二级表头）")
    @Test
    @Disabled
    public void writeSecHeadRowTest() {
        List<?> row1 = CollUtil.newArrayList(1, "aa", "bb", "cc", "dd", "ee");
        List<?> row2 = CollUtil.newArrayList(2, "aa1", "bb1", "cc1", "dd1", "ee1");
        List<?> row3 = CollUtil.newArrayList(3, "aa2", "bb2", "cc2", "dd2", "ee2");
        List<?> row4 = CollUtil.newArrayList(4, "aa3", "bb3", "cc3", "dd3", "ee3");
        List<?> row5 = CollUtil.newArrayList(5, "aa4", "bb4", "cc4", "dd4", "ee4");

        List<List<?>> rows = CollUtil.newArrayList(row1, row2, row3, row4, row5);

        // 通过工具类创建writer
        ExcelWriter writer = ExcelUtil.getWriter("d:/test/writeSecHeadRowTest.xlsx");

        CellStyle cellStyle = writer.getWorkbook().createCellStyle();
        cellStyle.setWrapText(false);
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        //设置标题内容字体
        Font font = writer.createFont();
        font.setBold(true);
        font.setFontHeightInPoints((short) 15);
        font.setFontName("Arial");
        //设置边框样式
        StyleUtil.setBorder(cellStyle, BorderStyle.THICK, IndexedColors.RED);
        cellStyle.setFont(font);

        // 合并单元格后的标题行，使用设置好的样式
        writer.merge(0, 1, 0, row1.size() - 1, "标题XXXXXXXX", cellStyle);
        Console.log(writer.getCurrentRow());

        //设置复杂表头
        writer.merge(2, 3, 0, 0, "序号", true);
        writer.merge(2, 2, 1, 2, "AABB", true);
        writer.merge(2, 3, 3, 3, "CCCC", true);
        writer.merge(2, 2, 4, 5, "DDEE", true);
        writer.setCurrentRow(3);

        List<String> sechead = CollUtil.newArrayList("AA", "BB", "DD", "EE");
        writer.writeSecHeadRow(sechead);
        // 一次性写出内容，使用默认样式
        writer.write(rows);
        // 关闭writer，释放内存
        writer.close();
    }

    /**
     * issue#1659@Github。
     * 测试使用BigWriter写出，ExcelWriter修改失败
     */
    @DisplayName("测试 BigExcelWriter 写出后用 ExcelWriter 修改文件")
    @Test
    @Disabled
    public void editTest() {
        // 生成文件
        File file = new File("d:/test/100_.xlsx");
        FileUtil.del(file);

        BigExcelWriter writer = ExcelUtil.getBigWriter(file);
        writer.disableDefaultStyle();
        List<List<String>> rows = Collections.singletonList(Arrays.asList("哈哈", "嘿嘿"));
        writer.write(rows);
        writer.close();

        // 修改文件
        ExcelWriter writer2 = ExcelUtil.getWriter(file);
        writer2.disableDefaultStyle();
        writer2.writeCellValue(0, 0, "a");
        writer2.close();

        final ExcelReader reader = ExcelUtil.getReader(file);
        Console.log(reader.read());
    }

    @DisplayName("测试合并单元格并写出富文本内容")
    @Test
    @Disabled
    public void mergeTest3() {
        // https://github.com/chinabugotech/hutool/issues/1696

        List<Map<String, Object>> list = new ArrayList<>();
        Map<String, Object> map = new HashMap<>();
        map.put("xmnf", "2021");
        list.add(map);

        Map<String, Object> map1 = new HashMap<>();
        map1.put("xmnf", new XSSFRichTextString("9999"));
        list.add(map1);

        Map<String, Object> map2 = new HashMap<>();
        map2.put("xmnf", "2019");
        list.add(map2);

        //通过工具类创建writer
        FileUtil.del("d:/test/writeTest2123.xlsx");
        ExcelWriter writer = ExcelUtil.getWriter("d:/test/writeTest2123.xlsx");
        writer.addHeaderAlias("xmnf", "项目年份"); //1

        //合并单元格后的标题行，使用默认标题样式
        writer.merge(7, "测试标题");
        writer.merge(3, 4, 0, 0, new XSSFRichTextString("9999"), true);
        writer.write(list, true);
        writer.close();
    }

    @DisplayName("测试合并单元格写入日期")
    @Test
    @Disabled
    public void mergeForDateTest() {
        // https://github.com/chinabugotech/hutool/issues/1911

        //通过工具类创建writer
        String path = "d:/test/mergeForDate.xlsx";
        FileUtil.del(path);
        ExcelWriter writer = ExcelUtil.getWriter(path);
        writer.merge(0, 3, 0, 2, DateUtil.date(), false);
        writer.close();
    }

    @DisplayName("测试修改表头单元格样式")
    @Test
    @Disabled
    public void changeHeaderStyleTest() {
        final ExcelWriter writer = ExcelUtil.getWriter("d:/test/headerStyle.xlsx");
        writer.writeHeadRow(ListUtil.of("姓名", "性别", "年龄"));
        final CellStyle headCellStyle = writer.getStyleSet().getHeadCellStyle();
        headCellStyle.setFillForegroundColor(IndexedColors.YELLOW1.index);
        headCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        writer.close();
    }

    @DisplayName("测试写出 Float 类型数据")
    @Test
    @Disabled
    public void writeFloatTest() {
        //issue https://gitee.com/chinabugotech/hutool/issues/I43U9G
        String path = "d:/test/floatTest.xlsx";
        FileUtil.del(path);

        final ExcelWriter writer = ExcelUtil.getWriter(path);
        writer.writeRow(ListUtil.of(22.9f));
        writer.close();
    }

    @DisplayName("测试写出带转义前缀的字符串单元格")
    @Test
    @Disabled
    public void issueI466ZZTest() {
        // https://gitee.com/chinabugotech/hutool/issues/I466ZZ
        // 需要输出S_20000314_x5116_0004
        // 此处加入一个转义前缀：_x005F
        List<Object> row = ListUtil.of(new EscapeStrCellSetter("S_20000314_x5116_0004"));

        ExcelWriter writer = ExcelUtil.getWriter("d:/test/_x.xlsx");
        writer.writeRow(row);
        writer.close();
    }

    @DisplayName("测试写出 Long 类型数据")
    @Test
    @Disabled
    public void writeLongTest() {
        //https://gitee.com/chinabugotech/hutool/issues/I49R6U
        final ExcelWriter writer = ExcelUtil.getWriter("d:/test/long.xlsx");
        writer.write(ListUtil.of(1427545395336093698L));
        writer.close();
    }

    @DisplayName("测试写出超链接单元格")
    @Test
    @Disabled
    public void writeHyperlinkTest() {
        final ExcelWriter writer = ExcelUtil.getWriter("d:/test/hyperlink.xlsx");

        final Hyperlink hyperlink = writer.createHyperlink(HyperlinkType.URL, "https://hutool.cn");

        writer.write(ListUtil.of(hyperlink));
        writer.close();
    }

    @DisplayName("测试合并单元格写入数字")
    @Test
    @Disabled
    public void mergeNumberTest() {
        File tempFile = new File("d:/test/mergeNumber.xlsx");
        FileUtil.del(tempFile);

        BigExcelWriter writer = new BigExcelWriter(tempFile);
        writer.merge(0, 1, 2, 2, 3.99, false);
        writer.close();
    }

    @DisplayName("测试写出图片到 Excel")
    @Test
    @Disabled
    public void writeImgTest() {
        ExcelWriter writer = ExcelUtil.getWriter(true);

        File file = new File("C:\\Users\\zsz\\Desktop\\1.jpg");

        writer.writeImg(file, 0, 0, 5, 10);

        writer.flush(new File("C:\\Users\\zsz\\Desktop\\2.xlsx"));

        writer.close();
    }

    @DisplayName("测试获取下载文件的 Content-Disposition 头信息")
    @Test
    public void getDispositionTest() {
        ExcelWriter writer = ExcelUtil.getWriter(true);
        final String disposition = writer.getDisposition("测试A12.xlsx", CharsetUtil.CHARSET_UTF_8);
        assertEquals("attachment; filename=\"%E6%B5%8B%E8%AF%95A12.xlsx\"", disposition);
    }

    @DisplayName("测试按列写出数据")
    @Test
    @Disabled
    public void writeColTest() {
        ExcelWriter writer = ExcelUtil.getWriter("G:\\test.xlsx");
        List<Integer> list = new ArrayList<>();
        for (int i = 0; i < 20; i++) {
            list.add(i);
        }
        writer.writeCol("header1", 0, list, true);
        writer.writeCol("header2", 1, list, false);
        writer.writeCol("header3", 0, list, true);
        writer.writeCol("header4", 1, list, false);

        Map<String, List<Integer>> map1 = new LinkedHashMap<>();
        map1.put("map1_header1", list);
        map1.put("map1_header2", list);
        map1.put("map1_header3", list);
        map1.put("map1_header4", list);

        Map<String, List<Integer>> map2 = new LinkedHashMap<>();
        map2.put("map2_header1", list);
        map2.put("map2_header2", list);
        map2.put("map2_header3", list);
        map2.put("map2_header4", list);

        writer.writeCol(map1, true);
        writer.writeCol(map2, false);

        writer.close();
    }

    @DisplayName("测试自动调整列宽")
    @Test
    @Disabled
    public void autoSizeColumnTest() {

        Map<String, Object> map = new LinkedHashMap<>(MAP_DATA_1);
        map.put("中文长度测试（符号）", "abc");

        String file1 = "d:/autoSizeColumnTest.xlsx";
        String file2 = "d:/autoSizeColumnTest2.xlsx";

        FileUtil.del(file1);
        FileUtil.del(file2);

        try (ExcelWriter writer = new ExcelWriter(file1)) {
            writer.writeRow(map, true);
            writer.autoSizeColumnAll(2f);
        }

        try (BigExcelWriter writer = new BigExcelWriter(file2)) {
            writer.writeRow(map, true);
            writer.autoSizeColumnAll(2f);
        }
    }

    @DisplayName("测试使用比较器指定列顺序写出大数据")
    @Test
    @Disabled
    public void writeWithComparatorTest() {
        // 生成测试数据, 10w行50列
        List<Map<String, Object>> dataList = new ArrayList<>();
        for (int i = 1; i <= 100000; i++) {
            Map<String, Object> map = new HashMap<>();
            map.put("test11", "test11_" + i);
            map.put("test12", "test12_" + i);
            map.put("test13", "test13_" + i);
            map.put("test14", "test14_" + i);
            map.put("test15", "test15_" + i);
            map.put("test16", "test16_" + i);
            map.put("test17", "test17_" + i);
            map.put("test18", "test18_" + i);
            map.put("test19", "test19_" + i);
            map.put("test1", "test1_" + i);
            map.put("test2", "test2_" + i);
            map.put("test3", "test3_" + i);
            map.put("test4", "test4_" + i);
            map.put("test5", "test5_" + i);
            map.put("test6", "test6_" + i);
            map.put("test7", "test7_" + i);
            map.put("test8", "test8_" + i);
            map.put("test9", "test9_" + i);
            map.put("test10", "test10_" + i);
            map.put("test20", "test20_" + i);
            map.put("test21", "test21_" + i);
            map.put("test22", "test22_" + i);
            map.put("test23", "test23_" + i);
            map.put("test24", "test24_" + i);
            map.put("test25", "test25_" + i);
            map.put("test26", "test26_" + i);
            map.put("test27", "test27_" + i);
            map.put("test28", "test28_" + i);
            map.put("test29", "test29_" + i);
            map.put("test30", "test30_" + i);
            map.put("test41", "test41_" + i);
            map.put("test42", "test42_" + i);
            map.put("test43", "test43_" + i);
            map.put("test44", "test44_" + i);
            map.put("test45", "test45_" + i);
            map.put("test46", "test46_" + i);
            map.put("test47", "test47_" + i);
            map.put("test48", "test48_" + i);
            map.put("test49", "test49_" + i);
            map.put("test50", "test50_" + i);
            map.put("test31", "test31_" + i);
            map.put("test32", "test32_" + i);
            map.put("test33", "test33_" + i);
            map.put("test34", "test34_" + i);
            map.put("test35", "test35_" + i);
            map.put("test36", "test36_" + i);
            map.put("test37", "test37_" + i);
            map.put("test38", "test38_" + i);
            map.put("test39", "test39_" + i);
            map.put("test40", "test40_" + i);
            dataList.add(map);
        }
        // 使用比较器写出
        try (BigExcelWriter excelWriter = ExcelUtil.getBigWriter("d:/writeWithComparatorTest.xlsx")) {
            excelWriter.write(dataList, new IndexedComparator<>(
                "test1", "test2", "test3", "test4", "test5", "test6", "test7", "test8", "test9", "test10",
                "test11", "test12", "test13", "test14", "test15", "test16", "test17", "test18", "test19", "test20",
                "test21", "test22", "test23", "test24", "test25", "test26", "test27", "test28", "test29", "test30",
                "test31", "test32", "test33", "test34", "test35", "test36", "test37", "test38", "test39", "test40",
                "test41", "test42", "test43", "test44", "test45", "test46", "test47", "test48", "test49", "test50"
            ));
        }
    }

}
