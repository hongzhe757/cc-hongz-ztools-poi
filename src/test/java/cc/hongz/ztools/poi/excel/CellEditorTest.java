package cc.hongz.ztools.poi.excel;

import cc.hongz.ztools.poi.excel.cell.CellEditor;
import lombok.AllArgsConstructor;
import lombok.Data;
import org.apache.poi.ss.usermodel.Cell;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import java.io.Serializable;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;

@DisplayName("CellEditor 单元格编辑器读取测试")
public class CellEditorTest {

    @Test
    @DisplayName("读取 Excel 时使用自定义 CellEditor 编辑单元格值")
    public void readTest() {
        ExcelReader excelReader = ExcelUtil.getReader("cell_editor_test.xlsx");
        excelReader.setCellEditor(new ExcelHandler());
        List<TestBean> excelReaderObjects = excelReader.readAll(TestBean.class);

        assertEquals("0", excelReaderObjects.get(0).getTest1());
        assertEquals("b", excelReaderObjects.get(0).getTest2());
        assertEquals("0", excelReaderObjects.get(1).getTest1());
        assertEquals("b1", excelReaderObjects.get(1).getTest2());
        assertEquals("0", excelReaderObjects.get(2).getTest1());
        assertEquals("c2", excelReaderObjects.get(2).getTest2());
    }

    @AllArgsConstructor
    @Data
    public static class TestBean implements Serializable {
        private static final long serialVersionUID = 1L;

        private String test1;
        private String test2;
    }

    public static class ExcelHandler implements CellEditor {
        @ Override
        public Object edit(final Cell cell, final Object o) {
            Object value = o;
            if (cell.getColumnIndex() == 0 && cell.getRowIndex() != 0) {
                value = "0";
            }
            return value;
        }
    }
}
