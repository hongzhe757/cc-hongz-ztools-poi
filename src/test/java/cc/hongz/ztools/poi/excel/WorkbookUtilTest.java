package cc.hongz.ztools.poi.excel;

import org.apache.poi.ss.usermodel.Workbook;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

@DisplayName("WorkbookUtil 创建工作簿测试")
public class WorkbookUtilTest {

    @Test
    @DisplayName("分别创建 xls 和 xlsx 工作簿并断言非空")
    public void createBookTest() {
        Workbook book = WorkbookUtil.createBook(true);
        assertNotNull(book);

        book = WorkbookUtil.createBook(false);
        assertNotNull(book);
    }
}
