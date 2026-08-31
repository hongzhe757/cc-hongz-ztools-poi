package cc.hongz.ztools.poi.excel;

import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.junit.jupiter.api.Disabled;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import cc.hongz.ztools.core.lang.Console;

@DisplayName("CellUtil 单元格工具类测试")
public class CellUtilTest {

    @Test
    @Disabled
    @DisplayName("遍历打印所有内置格式以校验日期格式")
    public void isDateTest() {
        String[] all = BuiltinFormats.getAll();
        for (int i = 0; i < all.length; i++) {
            Console.log("{} {}", i, all[i]);
        }
    }
}
