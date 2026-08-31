package cc.hongz.ztools.poi.excel;

import cc.hongz.ztools.core.io.IORuntimeException;
import cc.hongz.ztools.poi.excel.cell.CellUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.jupiter.api.Disabled;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;

/**
 * https://gitee.com/chinabugotech/hutool/issues/I6MBS5。 <br>
 * 经过测试，发现BigExcelWriter中的comment会错位<br>
 * 修正方式见: https://stackoverflow.com/questions/28169011/using-sxssfapache-poi-and-adding-comment-does-not-generate-proper-excel-file
 */
@DisplayName("BigExcelWriter单元格注释错位问题测试")
public class IssueI6MBS5Test {

    @DisplayName("BigExcelWriter设置单元格注释测试")
    @Test
    @Disabled
    public void setCommentTest() {
        final ExcelWriter writer = ExcelUtil.getBigWriter("d:/test/setCommentTest.xlsx");
        final Cell cell = writer.getOrCreateCell(0, 0);
        CellUtil.setCellValue(cell, "cellValue");
        CellUtil.setComment(cell, "commonText", "ascend", null);

        writer.close();
    }

    @DisplayName("CellUtil设置单元格注释并写出测试")
    @Test
    @Disabled
    public void setCommentTest2() {
        final File file = new File("D:\\test\\CellUtilTest.xlsx");
        try (final Workbook workbook = WorkbookUtil.createBook(true)) {
            final Sheet sheet = workbook.createSheet();
            final Row row = sheet.createRow(0);
            final Cell cell = row.createCell(0);
            CellUtil.setCellValue(cell, "cellValue");
            CellUtil.setComment(cell, "commonText", "ascend", null);
            workbook.write(Files.newOutputStream(file.toPath()));
        } catch (final IOException e) {
            throw new IORuntimeException(e);
        }
    }
}
