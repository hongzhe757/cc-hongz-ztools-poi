package cc.hongz.ztools.poi.excel;

import cc.hongz.ztools.core.lang.Console;

import org.junit.jupiter.api.Disabled;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

/**
 * 测试使用ExcelUtil.readBySax读取.xls文件时，公式结果为字符串的单元格内容不缺失。
 *
 * @see <a href = "https://gitee.com/chinabugotech/hutool/issues/IJSHQT">Issue IJSHQT</a>
 */
@DisplayName("issue IJSHQT：Sax 读取 .xls 公式字符串单元格不缺失")
public class IssueIJSHQTTest {

    @Test
    @Disabled
    @DisplayName("Sax 方式读取 .xls 校验公式字符串单元格内容")
    public void readBySaxXlsTest() {
        // 替换为你的xls文件路径
        ExcelUtil.readBySax("d:/test/b106e21a-78e1-49a7-86e4-da70035ac6ed.xls", 0,
                (sheetIndex, rowIndex, rowList) -> {
                    Console.log("[{}] [{}] {}", sheetIndex, rowIndex, rowList);
                });
    }
}
