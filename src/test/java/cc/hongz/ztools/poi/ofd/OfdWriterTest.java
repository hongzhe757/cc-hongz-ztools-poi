package cc.hongz.ztools.poi.ofd;

import cc.hongz.ztools.core.io.FileUtil;
import org.junit.jupiter.api.Disabled;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

@DisplayName("OfdWriter 写出 OFD 文件测试")
public class OfdWriterTest {

    @Test
    @Disabled
    @DisplayName("写出 OFD 文件并添加文本内容")
    public void writeTest() {
        final OfdWriter ofdWriter = new OfdWriter(FileUtil.file("d:/test/test.ofd"));
        ofdWriter.addText(null, "测试文本");
        ofdWriter.close();
    }
}
