package cc.hongz.ztools.poi.excel.cell.setters;

import cc.hongz.ztools.core.util.StrUtil;
import cc.hongz.ztools.poi.excel.cell.CellSetter;
import org.apache.poi.ss.usermodel.Cell;

/**
 * {@link Number} 值单元格设置器。
 *
 * @author looly
 * @since 5.7.8
 */
public class NullCellSetter implements CellSetter {

    /**
     * 单例实例。
     */
    public static final NullCellSetter INSTANCE = new NullCellSetter();

    @Override
    public void setValue(final Cell cell) {
        cell.setCellValue(StrUtil.EMPTY);
    }
}
