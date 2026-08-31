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

import cc.hongz.ztools.core.io.FileUtil;
import org.junit.jupiter.api.Disabled;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

@DisplayName("ExcelReader 获取 Writer 写出测试")
public class ExcelReaderToWriterTest {

    @Test
    @Disabled
    @DisplayName("从 ExcelReader 获取 Writer 并写出替换内容")
    public void getWriterTest() {
        final ExcelReader reader = ExcelUtil.getReader("d:/test/issue3120.xlsx");
        final ExcelWriter writer = reader.getWriter();
        writer.writeCellValue(0, 0, "替换内容222222");

        writer.flush(FileUtil.file("d:/test/2.xlsx"));
        writer.closeWithoutFlush();
    }
}
