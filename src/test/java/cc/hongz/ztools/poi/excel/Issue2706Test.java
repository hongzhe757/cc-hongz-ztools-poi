package cc.hongz.ztools.poi.excel;

import org.junit.jupiter.api.Disabled;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

@DisplayName("issue#2706：BigExcelWriter 写出 Map 数据")
public class Issue2706Test {

    @Test
    @Disabled
    @DisplayName("使用 BigExcelWriter 写出 Map 数据列表")
    public void writeTest() {
        String path = "d:/test/issue2706.xlsx";
        BigExcelWriter writer = ExcelUtil.getBigWriter(path, "表格名称");

        Set<String> headSet1 = new HashSet<>(Arrays.asList("A", "B", "C", "D", "E"));

        List<Map<String, String>> datas = new ArrayList<>();
        for (int i = 0; i < 10; i++) {
            Map<String, String> map = new HashMap<>();
            for (String s : headSet1) {
                map.put(s, "h" + i);
            }
            datas.add(map);
        }
        writer.write(datas);
        writer.close();
    }
}
