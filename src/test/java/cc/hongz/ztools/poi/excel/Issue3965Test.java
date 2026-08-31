package cc.hongz.ztools.poi.excel;

import cc.hongz.ztools.core.collection.ListUtil;
import cc.hongz.ztools.core.io.FileUtil;
import org.junit.jupiter.api.Disabled;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import java.util.ArrayList;
import java.util.List;

@DisplayName("issue#3965：Excel 写出 List 数据并设置列宽")
public class Issue3965Test {
    @Test
    @Disabled
    @DisplayName("写出嵌套 List 数据并设置第一列列宽")
    void writeTest() {
        ArrayList<List<String>> arrayList = new ArrayList<>();

        arrayList.add(ListUtil.of("a"));
        arrayList.add(ListUtil.of("b"));
        arrayList.add(ListUtil.of("c"));
        arrayList.add(ListUtil.of("d"));

        ExcelWriter writer = ExcelUtil.getWriter(FileUtil.file("d:/test/123.xlsx"));
        writer.setColumnWidth(0, 50);
        writer.write(arrayList);
        writer.flush();
        writer.close();
    }
}
