package cc.hongz.ztools.poi.excel.cell;

import cc.hongz.ztools.core.util.ObjectUtil;
import cc.hongz.ztools.core.util.StrUtil;
import cc.hongz.ztools.poi.excel.ExcelUtil;
import cc.hongz.ztools.poi.excel.StyleSet;
import cc.hongz.ztools.poi.excel.cell.setters.CellSetterFactory;
import cc.hongz.ztools.poi.excel.cell.values.ErrorCellValue;
import cc.hongz.ztools.poi.excel.cell.values.NumericCellValue;
import cc.hongz.ztools.poi.excel.editors.TrimEditor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.ss.util.SheetUtil;

/**
 * Excel表格中单元格工具类。
 *
 * @author looly
 * @since 4.0.7
 */
@SuppressWarnings("checkstyle:HideUtilityClassConstructor")
public class CellUtil {

    /**
     * 获取单元格值。
     *
     * @param cell {@link Cell}单元格
     * @return 值，类型可能为：Date、Double、Boolean、String
     * @since 4.6.3
     */
    public static Object getCellValue(final Cell cell) {
        return getCellValue(cell, false);
    }

    /**
     * 获取单元格值。
     *
     * @param cell            {@link Cell}单元格
     * @param isTrimCellValue 如果单元格类型为字符串，是否去掉两边空白符
     * @return 值，类型可能为：Date、Double、Boolean、String
     */
    public static Object getCellValue(final Cell cell, final boolean isTrimCellValue) {
        if (null == cell) {
            return null;
        }
        return getCellValue(cell, cell.getCellType(), isTrimCellValue);
    }

    /**
     * 获取单元格值。
     *
     * @param cell       {@link Cell}单元格
     * @param cellEditor 单元格值编辑器。可以通过此编辑器对单元格值做自定义操作
     * @return 值，类型可能为：Date、Double、Boolean、String
     */
    public static Object getCellValue(final Cell cell, final CellEditor cellEditor) {
        return getCellValue(cell, null, cellEditor);
    }

    /**
     * 获取单元格值。
     *
     * @param cell            {@link Cell}单元格
     * @param cellType        单元格值类型{@link CellType}枚举
     * @param isTrimCellValue 如果单元格类型为字符串，是否去掉两边空白符
     * @return 值，类型可能为：Date、Double、Boolean、String
     */
    public static Object getCellValue(final Cell cell, final CellType cellType, final boolean isTrimCellValue) {
        return getCellValue(cell, cellType, isTrimCellValue ? new TrimEditor() : null);
    }

    /**
     * 获取单元格值。 <br>
     * 如果单元格值为数字格式，则判断其格式中是否有小数部分，无则返回Long类型，否则返回Double类型
     *
     * @param cell       {@link Cell}单元格
     * @param cellType   单元格值类型{@link CellType}枚举，如果为{@code null}默认使用cell的类型
     * @param cellEditor 单元格值编辑器。可以通过此编辑器对单元格值做自定义操作
     * @return 值，类型可能为：Date、Double、Boolean、String
     */
    public static Object getCellValue(final Cell cell, final CellType cellType, final CellEditor cellEditor) {
        if (null == cell) {
            return null;
        }
        if (cell instanceof NullCell) {
            return null == cellEditor ? null : cellEditor.edit(cell, null);
        }
        Cell resolvedCell = cell;
        CellType resolvedCellType = cellType;
        if (null == resolvedCellType) {
            resolvedCellType = resolvedCell.getCellType();
        }

        // 尝试获取合并单元格，如果是合并单元格，则重新获取单元格类型
        final Cell mergedCell = getMergedRegionCell(resolvedCell);
        if (mergedCell != resolvedCell) {
            resolvedCell = mergedCell;
            resolvedCellType = resolvedCell.getCellType();
        }

        final Object value;
        switch (resolvedCellType) {
            case NUMERIC:
                value = new NumericCellValue(resolvedCell).getValue();
                break;
            case BOOLEAN:
                value = resolvedCell.getBooleanCellValue();
                break;
            case FORMULA:
                value = getCellValue(resolvedCell, resolvedCell.getCachedFormulaResultType(), cellEditor);
                break;
            case BLANK:
                value = StrUtil.EMPTY;
                break;
            case ERROR:
                value = new ErrorCellValue(resolvedCell).getValue();
                break;
            default:
                value = resolvedCell.getStringCellValue();
        }

        return null == cellEditor ? value : cellEditor.edit(resolvedCell, value);
    }

    /**
     * 设置单元格值。 <br>
     * 根据传入的styleSet自动匹配样式<br>
     * 当为头部样式时默认赋值头部样式，但是头部中如果有数字、日期等类型，将按照数字、日期样式设置
     *
     * @param cell     单元格
     * @param value    值
     * @param styleSet 单元格样式集，包括日期等样式，null表示无样式
     * @param isHeader 是否为标题单元格
     */
    public static void setCellValue(final Cell cell, final Object value, final StyleSet styleSet,
            final boolean isHeader) {
        if (null == cell) {
            return;
        }

        if (null != styleSet) {
            cell.setCellStyle(styleSet.getStyleByValueType(value, isHeader));
        }

        setCellValue(cell, value);
    }

    /**
     * 设置单元格值。 <br>
     * 根据传入的styleSet自动匹配样式<br>
     * 当为头部样式时默认赋值头部样式，但是头部中如果有数字、日期等类型，将按照数字、日期样式设置
     *
     * @param cell  单元格
     * @param value 值
     * @param style 自定义样式，null表示无样式
     */
    public static void setCellValue(final Cell cell, final Object value, final CellStyle style) {
        setCellValue(cell, (CellSetter) cell1 -> {
            setCellValue(cell, value);
            if (null != style) {
                cell1.setCellStyle(style);
            }
        });
    }

    /**
     * 设置单元格值。 <br>
     * 根据传入的styleSet自动匹配样式<br>
     * 当为头部样式时默认赋值头部样式，但是头部中如果有数字、日期等类型，将按照数字、日期样式设置
     *
     * @param cell  单元格
     * @param value 值或{@link CellSetter}
     * @since 5.6.4
     */
    public static void setCellValue(final Cell cell, final Object value) {
        if (null == cell) {
            return;
        }

        // issue#1659@Github
        // 在使用BigWriter(SXSSF)模式写出数据时，单元格值为直接值，非引用值（is标签）
        // 而再使用ExcelWriter(XSSF)编辑时，会写出引用值，导致失效。
        // 此处做法是先清空单元格值，再写入
        if (CellType.BLANK != cell.getCellType()) {
            cell.setBlank();
        }

        CellSetterFactory.createCellSetter(value).setValue(cell);
    }

    /**
     * 获取单元格，如果单元格不存在，返回{@link NullCell}。
     *
     * @param row       Excel表的行
     * @param cellIndex 列号
     * @return {@link Cell}
     * @since 5.5.0
     */
    public static Cell getCell(final Row row, final int cellIndex) {
        if (null == row) {
            return null;
        }
        final Cell cell = row.getCell(cellIndex);
        if (null == cell) {
            return new NullCell(row, cellIndex);
        }
        return cell;
    }

    /**
     * 获取已有单元格或创建新单元格。
     *
     * @param row       Excel表的行
     * @param cellIndex 列号
     * @return {@link Cell}
     * @since 4.0.2
     */
    public static Cell getOrCreateCell(final Row row, final int cellIndex) {
        if (null == row) {
            return null;
        }
        Cell cell = row.getCell(cellIndex);
        if (null == cell) {
            cell = row.createCell(cellIndex);
        }
        return cell;
    }

    /**
     * 判断指定的单元格是否是合并单元格。
     *
     * @param sheet       {@link Sheet}
     * @param locationRef 单元格地址标识符，例如A11，B5
     * @return 是否是合并单元格
     * @since 5.1.5
     */
    public static boolean isMergedRegion(final Sheet sheet, final String locationRef) {
        final CellLocation cellLocation = ExcelUtil.toLocation(locationRef);
        return isMergedRegion(sheet, cellLocation.getX(), cellLocation.getY());
    }

    /**
     * 判断指定的单元格是否是合并单元格。
     *
     * @param cell {@link Cell}
     * @return 是否是合并单元格
     * @since 5.1.5
     */
    public static boolean isMergedRegion(final Cell cell) {
        return isMergedRegion(cell.getSheet(), cell.getColumnIndex(), cell.getRowIndex());
    }

    /**
     * 判断指定的单元格是否是合并单元格。
     *
     * @param sheet {@link Sheet}
     * @param x     列号，从0开始
     * @param y     行号，从0开始
     * @return 是否是合并单元格，如果提供的sheet为{@code null}，返回{@code false}
     */
    public static boolean isMergedRegion(final Sheet sheet, final int x, final int y) {
        if (sheet != null) {
            final int sheetMergeCount = sheet.getNumMergedRegions();
            CellRangeAddress ca;
            for (int i = 0; i < sheetMergeCount; i++) {
                ca = sheet.getMergedRegion(i);
                if (y >= ca.getFirstRow() && y <= ca.getLastRow()
                        && x >= ca.getFirstColumn() && x <= ca.getLastColumn()) {
                    return true;
                }
            }
        }
        return false;
    }

    /**
     * 获取合并单元格{@link CellRangeAddress}，如果不是返回null。
     *
     * @param sheet       {@link Sheet}
     * @param locationRef 单元格地址标识符，例如A11，B5
     * @return {@link CellRangeAddress}
     * @since 5.8.0
     */
    public static CellRangeAddress getCellRangeAddress(final Sheet sheet, final String locationRef) {
        final CellLocation cellLocation = ExcelUtil.toLocation(locationRef);
        return getCellRangeAddress(sheet, cellLocation.getX(), cellLocation.getY());
    }

    /**
     * 获取合并单元格{@link CellRangeAddress}，如果不是返回null。
     *
     * @param cell {@link Cell}
     * @return {@link CellRangeAddress}
     * @since 5.8.0
     */
    public static CellRangeAddress getCellRangeAddress(final Cell cell) {
        return getCellRangeAddress(cell.getSheet(), cell.getColumnIndex(), cell.getRowIndex());
    }

    /**
     * 获取合并单元格{@link CellRangeAddress}，如果不是返回null。
     *
     * @param sheet {@link Sheet}
     * @param x     列号，从0开始
     * @param y     行号，从0开始
     * @return {@link CellRangeAddress}
     * @since 5.8.0
     */
    public static CellRangeAddress getCellRangeAddress(final Sheet sheet, final int x, final int y) {
        if (sheet != null) {
            final int sheetMergeCount = sheet.getNumMergedRegions();
            CellRangeAddress ca;
            for (int i = 0; i < sheetMergeCount; i++) {
                ca = sheet.getMergedRegion(i);
                if (y >= ca.getFirstRow() && y <= ca.getLastRow()
                        && x >= ca.getFirstColumn() && x <= ca.getLastColumn()) {
                    return ca;
                }
            }
        }
        return null;
    }

    /**
     * 设置合并单元格样式，如果不是则不设置。
     *
     * @param cell      {@link Cell}
     * @param cellStyle {@link CellStyle}
     */
    public static void setMergedRegionStyle(final Cell cell, final CellStyle cellStyle) {
        final CellRangeAddress cellRangeAddress = getCellRangeAddress(cell);
        if (cellRangeAddress != null) {
            setMergeCellStyle(cellStyle, cellRangeAddress, cell.getSheet());
        }
    }

    /**
     * 合并单元格，可以根据设置的值来合并行和列。
     *
     * @param sheet       表对象
     * @param firstRow    起始行，0开始
     * @param lastRow     结束行，0开始
     * @param firstColumn 起始列，0开始
     * @param lastColumn  结束列，0开始
     * @return 合并后的单元格号
     */
    public static int mergingCells(final Sheet sheet, final int firstRow, final int lastRow,
            final int firstColumn, final int lastColumn) {
        return mergingCells(sheet, firstRow, lastRow, firstColumn, lastColumn, null);
    }

    /**
     * 合并单元格，可以根据设置的值来合并行和列。
     *
     * @param sheet       表对象
     * @param firstRow    起始行，0开始
     * @param lastRow     结束行，0开始
     * @param firstColumn 起始列，0开始
     * @param lastColumn  结束列，0开始
     * @param cellStyle   单元格样式，只提取边框样式，null表示无样式
     * @return 合并后的单元格号
     */
    public static int mergingCells(final Sheet sheet, final int firstRow, final int lastRow,
            final int firstColumn, final int lastColumn, final CellStyle cellStyle) {
        final CellRangeAddress cellRangeAddress = new CellRangeAddress(//
                firstRow, // first row (0-based)
                lastRow, // last row (0-based)
                firstColumn, // first column (0-based)
                lastColumn // last column (0-based)
        );

        setMergeCellStyle(cellStyle, cellRangeAddress, sheet);
        return sheet.addMergedRegion(cellRangeAddress);
    }

    /**
     * 获取合并单元格的值。 <br>
     * 传入的x,y坐标（列行数）可以是合并单元格范围内的任意一个单元格
     *
     * @param sheet       {@link Sheet}
     * @param locationRef 单元格地址标识符，例如A11，B5
     * @return 合并单元格的值
     * @since 5.1.5
     */
    public static Object getMergedRegionValue(final Sheet sheet, final String locationRef) {
        final CellLocation cellLocation = ExcelUtil.toLocation(locationRef);
        return getMergedRegionValue(sheet, cellLocation.getX(), cellLocation.getY());
    }

    /**
     * 获取合并单元格的值。 <br>
     * 传入的x,y坐标（列行数）可以是合并单元格范围内的任意一个单元格
     *
     * @param sheet {@link Sheet}
     * @param x     列号，从0开始，可以是合并单元格范围中的任意一列
     * @param y     行号，从0开始，可以是合并单元格范围中的任意一行
     * @return 合并单元格的值
     * @since 4.6.3
     */
    public static Object getMergedRegionValue(final Sheet sheet, final int x, final int y) {
        // 合并单元格的识别在getCellValue已经集成，无需重复获取合并单元格
        return getCellValue(SheetUtil.getCell(sheet, x, y));
    }

    /**
     * 获取合并单元格。 <br>
     * 传入的x,y坐标（列行数）可以是合并单元格范围内的任意一个单元格
     *
     * @param cell {@link Cell}
     * @return 合并单元格
     * @since 5.1.5
     */
    public static Cell getMergedRegionCell(final Cell cell) {
        if (null == cell) {
            return null;
        }
        return ObjectUtil.defaultIfNull(
                getCellIfMergedRegion(cell.getSheet(), cell.getColumnIndex(), cell.getRowIndex()),
                cell);
    }

    /**
     * 获取合并单元格。 <br>
     * 传入的x,y坐标（列行数）可以是合并单元格范围内的任意一个单元格
     *
     * @param sheet {@link Sheet}
     * @param x     列号，从0开始，可以是合并单元格范围中的任意一列
     * @param y     行号，从0开始，可以是合并单元格范围中的任意一行
     * @return 合并单元格，如果非合并单元格，返回坐标对应的单元格
     * @since 5.1.5
     */
    public static Cell getMergedRegionCell(final Sheet sheet, final int x, final int y) {
        return ObjectUtil.defaultIfNull(
                getCellIfMergedRegion(sheet, x, y),
                () -> SheetUtil.getCell(sheet, y, x));
    }

    /**
     * 为特定单元格添加批注。
     *
     * @param cell          单元格
     * @param commentText   批注内容
     * @param commentAuthor 作者
     * @param anchor        批注的位置、大小等信息，null表示使用默认
     * @since 5.4.8
     */
    public static void setComment(final Cell cell, final String commentText, final String commentAuthor,
            final ClientAnchor anchor) {
        final Sheet sheet = cell.getSheet();
        final Workbook wb = sheet.getWorkbook();
        final Drawing<?> drawing = sheet.createDrawingPatriarch();
        final CreationHelper factory = wb.getCreationHelper();
        ClientAnchor resolvedAnchor = anchor;
        if (resolvedAnchor == null) {
            resolvedAnchor = factory.createClientAnchor();
            // 默认位置，在注释的单元格的右方
            resolvedAnchor.setCol1(cell.getColumnIndex() + 1);
            resolvedAnchor.setCol2(cell.getColumnIndex() + 3);
            resolvedAnchor.setRow1(cell.getRowIndex());
            resolvedAnchor.setRow2(cell.getRowIndex() + 2);
            // 自适应
            resolvedAnchor.setAnchorType(ClientAnchor.AnchorType.MOVE_AND_RESIZE);
        }
        final Comment comment = drawing.createCellComment(resolvedAnchor);
        // https://stackoverflow.com/questions/28169011/using-sxssfapache-poi-and-adding-comment-does-not-generate-proper-excel-file
        // 修正在XSSFCell中未设置地址导致错位问题
        comment.setAddress(cell.getAddress());
        comment.setString(factory.createRichTextString(commentText));
        comment.setAuthor(StrUtil.nullToEmpty(commentAuthor));
        cell.setCellComment(comment);
    }

    // ---------------------------------------- Private method start

    /**
     * 获取合并单元格，非合并单元格返回{@code null}。 <br>
     * 传入的x,y坐标（列行数）可以是合并单元格范围内的任意一个单元格
     *
     * @param sheet {@link Sheet}
     * @param x     列号，从0开始，可以是合并单元格范围中的任意一列
     * @param y     行号，从0开始，可以是合并单元格范围中的任意一行
     * @return 合并单元格，如果非合并单元格，返回{@code null}
     * @since 5.4.5
     */
    private static Cell getCellIfMergedRegion(final Sheet sheet, final int x, final int y) {
        for (final CellRangeAddress ca : sheet.getMergedRegions()) {
            if (ca.isInRange(y, x)) {
                return SheetUtil.getCell(sheet, ca.getFirstRow(), ca.getFirstColumn());
            }
        }
        return null;
    }

    /**
     * 根据{@link CellStyle}设置合并单元格边框样式。
     *
     * @param cellStyle        {@link CellStyle}
     * @param cellRangeAddress {@link CellRangeAddress}
     * @param sheet            {@link Sheet}
     */
    private static void setMergeCellStyle(final CellStyle cellStyle, final CellRangeAddress cellRangeAddress,
            final Sheet sheet) {
        if (null != cellStyle) {
            RegionUtil.setBorderTop(cellStyle.getBorderTop(), cellRangeAddress, sheet);
            RegionUtil.setBorderRight(cellStyle.getBorderRight(), cellRangeAddress, sheet);
            RegionUtil.setBorderBottom(cellStyle.getBorderBottom(), cellRangeAddress, sheet);
            RegionUtil.setBorderLeft(cellStyle.getBorderLeft(), cellRangeAddress, sheet);
            RegionUtil.setTopBorderColor(cellStyle.getTopBorderColor(), cellRangeAddress, sheet);
            RegionUtil.setRightBorderColor(cellStyle.getRightBorderColor(), cellRangeAddress, sheet);
            RegionUtil.setLeftBorderColor(cellStyle.getLeftBorderColor(), cellRangeAddress, sheet);
            RegionUtil.setBottomBorderColor(cellStyle.getBottomBorderColor(), cellRangeAddress, sheet);
        }
    }
    // ---------------------------------------- Private method end
}
