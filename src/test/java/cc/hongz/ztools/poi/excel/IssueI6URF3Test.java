/*
 * Copyright (c) 2023 looly(loolly@aliyun.com)
 * Hutool is licensed under Mulan PSL v2.
 * You can use this software according to the terms and conditions of the Mulan PSL v2.
 * You may obtain a copy of Mulan PSL v2 at:
 *          http://license.coscl.org.cn/MulanPSL2
 * THIS SOFTWARE IS PROVIDED ON AN "AS IS" BASIS, WITHOUT WARRANTIES OF ANY KIND,
 * EITHER EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO NON-INFRINGEMENT,
 * MERCHANTABILITY OR FIT FOR A PARTICULAR PURPOSE.
 * See the Mulan PSL v2 for more details.
 */

package cc.hongz.ztools.poi.excel;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.junit.jupiter.api.Disabled;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

@DisplayName("issue I6URF3：设置单元格样式")
public class IssueI6URF3Test {

    @Test
    @Disabled
    @DisplayName("写入单元格并为其创建样式设置对齐方式")
    public void setCellStyleTest() {
        final ExcelWriter writer = ExcelUtil.getWriter("d:/test/issueI6URF3.xlsx");
        writer.writeCellValue(0, 0, 1);
        final CellStyle cellStyle = writer.createCellStyle(0, 0);
        cellStyle.setAlignment(HorizontalAlignment.LEFT);
        writer.close();
    }
}
