package com.lixi.util.excel;

import org.junit.Test;

import java.io.IOException;
import java.io.InputStream;
import java.util.List;
import java.util.Map;

import static org.junit.Assert.assertSame;

/**
 * Created by lixi on 17/8/14.
 */
public class ExcelUtilTest {

    @Test
    public void testExcelToList() throws IOException {
        InputStream input = ExcelUtilTest.class.getResourceAsStream("/testExcelToList.xlsx");
        List<Map<String, Object>> result = ExcelUtil.excelToList(input);
        assertSame(3, result.size());
    }
}
