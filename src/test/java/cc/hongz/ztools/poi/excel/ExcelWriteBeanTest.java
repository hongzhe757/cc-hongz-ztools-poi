package cc.hongz.ztools.poi.excel;

import lombok.Getter;
import org.junit.jupiter.api.Disabled;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

@DisplayName("ExcelWriter 写出 JavaBean 测试")
public class ExcelWriteBeanTest {

    @Test
    @Disabled
    @DisplayName("写出 JavaBean 作为一行数据")
    public void writeRowTest() {
        MyBean bean = new MyBean("value1", "value2");

        final ExcelWriter writer = ExcelUtil.getWriter("d:/test/writeRowTest.xlsx");
        writer.writeRow(bean, true);
        writer.close();
    }

    @Getter
    static class MyBean {
        private final String property1;
        private final String property2;

        public MyBean(final String property1, final String property2) {
            this.property1 = property1;
            this.property2 = property2;
        }
    }
}
