package cc.hongz.ztools.poi.excel;

import cc.hongz.ztools.core.annotation.Alias;
import cc.hongz.ztools.core.collection.CollUtil;
import cc.hongz.ztools.core.io.FileUtil;
import lombok.Data;
import org.junit.jupiter.api.Disabled;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

@DisplayName("issue#2899：@Alias 注解与 setOnlyAlias 写出")
public class Issue2899Test {

    @Test
    @Disabled
    @DisplayName("写出时校验 @Alias 注解与 setOnlyAlias 的生效关系")
    public void aliasWriteTest() {
        // Bean中设置@Alias时，setOnlyAlias是无效的，这个参数只和addHeaderAlias配合使用，原因是注解是Bean内部的操作，而addHeaderAlias是Writer的操作，不互通。
        final TestBean testBean1 = new TestBean();
        testBean1.setName("张三");
        testBean1.setScore(12);

        final TestBean testBean2 = new TestBean();
        testBean2.setName("李四");
        testBean2.setScore(23);

        FileUtil.del("d:/test/aliasTest.xlsx");
        final ExcelWriter writer = ExcelUtil.getWriter("d:/test/aliasTest.xlsx");

        writer.addHeaderAlias("姓名", "姓名");
        writer.setOnlyAlias(true);
        writer.merge(2, "成绩单");
        writer.write(CollUtil.newArrayList(testBean1, testBean2), true);
        writer.close();
    }

    @Data
    static class TestBean {
        @Alias("姓名")
        private String name;
        private double score;
    }
}
