
package com.lixi.util.excel;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import javax.script.ScriptEngine;
import javax.script.ScriptEngineManager;
import javax.script.ScriptException;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;


public class ExcelTemplateUtil {

    public static final String DEFAULT_SOURCE_NAME = "main";
    public static final String SOURCENAME_REG = "\\$s\\{([^\\}]*)\\}";
    private static final String MERGE_TAG = "$m";
    public static ScriptEngineManager mgr = new ScriptEngineManager();
    public static ScriptEngine engin = mgr.getEngineByName("JavaScript");

    private void fillField(HSSFSheet sheet,
            Map<String, Object> mapCondition) {
        try {
            int rownum = sheet.getLastRowNum();
            for (int i = 0; i < rownum + 1; i++) {
                HSSFRow row = sheet.getRow(i);
                if(row == null){
                    continue;
                }
                int cells = row.getLastCellNum();
                for (int j = 0; j < cells; j++) {
                    HSSFCell cell = row.getCell(j);
                    if (cell == null) {
                        continue;
                    }

                    String stringCellValue = cell.toString();
                    if (isField(stringCellValue)) {
                        Object value = getFieldJsValue(stringCellValue, mapCondition);
                        if (value instanceof Date) {
                            value = new SimpleDateFormat("yyyy-MM-dd")
                                    .format((Date) value);
                        }
                        if (value != null && isNumber(value)) {
                            cell.setCellValue(Double
                                    .parseDouble(value.toString()));
                        } else {
                            cell.setCellValue(new HSSFRichTextString(changeToString(value)));
                        }
                    }
                }
            }
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    /**
     * 将若干个数据源，根据Excel模板填充到Excel中
     * @param data 数据源列表
     * 格式{数据源名称,[数据列表]},如数据源中存放2个数据项Student, Course,其中每个数据项又可以包含List
     * {
     *  "Student" : [{"name":"张三", "sex":"男", "age":15, "address":[{"name" : "上海浦东"},{}]},{"name":"李芳", "sex":"女","age":18}],
     *  "Course" :  [{"name":"语文", "score":5},{"name":"数学", "score":10}]
     * }
     * @param templateStream 模板文件
     * @param fields 字段列表
     * @return 返回生成Excel的数据源
     */
    public InputStream fillData(Map<String,List<Map<String, Object>>> data, InputStream templateStream,
            Map<String, Object> fields){
        try {
            HSSFWorkbook book = new HSSFWorkbook(templateStream);
            int sheetNum = book.getNumberOfSheets();
            for(int i = 0; i < sheetNum ; i ++){
               HSSFSheet sheet = book.getSheetAt(i);
               fillSource(data, sheet, fields);
            }
            ByteArrayOutputStream output = new ByteArrayOutputStream();
            updateFormula(book);
            book.write(output);
            return  new ByteArrayInputStream(output.toByteArray());
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    private void fillSource(Map<String,List<Map<String, Object>>> data, HSSFSheet sheet,
    		Map<String, Object> mapCondition) throws ScriptException {
            fillField(sheet, mapCondition);
            HSSFRow templateRow = null;
            HSSFCell templateCell = null;
            int startRow = 0;
            while((templateCell = searchTemplateRow(sheet, startRow)) != null){
                String value = templateCell.getStringCellValue();
                templateRow = templateCell.getRow();
                String sourceName = getSourceName(value);
                if(sourceName == null){
                    sourceName = DEFAULT_SOURCE_NAME;
                }
                replaceCellSourceName(templateCell);
                List<CellRangeAddress> regins = getTemplateRowRegins(sheet, templateRow);
                List<Map<String, Object>> sourceData = getSourceData(sourceName, data);
                if(sourceData != null){
                 // 从第二行插入数据，模板行最后再插入
                    for (int i = 1; i < sourceData.size(); i++) {
                        Map<String, Object> map = sourceData.get(i);
                        insertLine(sheet, templateRow, i, map, true, regins);
                        startRow = templateRow.getRowNum() + i;
                    }
                }
                Map<String, Object> map = sourceData!=null && sourceData.size() > 0 ? sourceData.get(0)
                        : new HashMap<String, Object>();
                // 覆盖模板行
                HSSFRow trow = insertLine(sheet, templateRow, 0, map, false, null);
                mergeCellIfNeed(trow, startRow);
            }
        }

    private List<CellRangeAddress> getTemplateRowRegins(HSSFSheet sheet, HSSFRow templateRow){
        List<CellRangeAddress > regions = new ArrayList<CellRangeAddress>();
        int nums = sheet.getNumMergedRegions();
        for(int i = 0; i < nums; i ++){
            CellRangeAddress region = sheet.getMergedRegion(i);
            if(region.getFirstRow() == region.getLastRow()&& region.getLastRow() == templateRow.getRowNum()){
                regions.add(region);
            }
        }
        return regions;

    }
    private void replaceCellSourceName(HSSFCell templateCell){
        String value = templateCell.getStringCellValue();
        StringTemplate tem = new StringTemplate(value,SOURCENAME_REG);
        for (String name : tem.getVariables()) {
            tem.setVariable(name, "");
        }
        templateCell.setCellValue(tem.getParseResult());
    }

    private List<Map<String,Object>> getSourceData(String sourceName, Map<String,List<Map<String, Object>>> data){
        List<Map<String, Object>> list = data.get(sourceName);
        if(null != list){
            return list;
        }
        String[] nameList = sourceName.split("\\.");
        if(nameList.length != 2){
            return null;
        }
        List<Map<String, Object>>  dataList = data.get(nameList[0]);
        if(dataList.size() == 0){
            return null;
        }
        Map<String, Object> record = dataList.get(0);
        return (List) record.get(nameList[1]);
    }

    private HSSFCell searchTemplateRow(HSSFSheet sheet, int startRow){
        int rownum = sheet.getLastRowNum();
        for (int i = startRow; i < rownum + 1; i++) {
             HSSFRow row = sheet.getRow(i);
             if(row == null){
                 continue;
             }
             int cells = row.getLastCellNum();
             for (int j = 0; j < cells; j++) {
                 HSSFCell cell = row.getCell(j);
                 if (cell == null) {
                     continue;
                 }
                 String stringCellValue = cell.toString();
                 if (isData(stringCellValue)) {
                     return cell;
                 }
             }
         }
        return null;

    }
    private HSSFRow insertLine(HSSFSheet sheet, HSSFRow templateRow, int i,
            Map<String, Object> map, boolean formatCondition, List<CellRangeAddress> regins) throws ScriptException {
        int num = templateRow.getRowNum();
        if(formatCondition && num + i <= sheet.getLastRowNum()){
            sheet.shiftRows(num + i, sheet.getLastRowNum(), 1, true, false);
        }
        if(regins != null && regins.size() > 0){
            for(CellRangeAddress region : regins){
                CellRangeAddress copyRegin = region.copy();
                copyRegin.setFirstRow(num + i);
                copyRegin.setLastRow(num + i);
                sheet.addMergedRegion(copyRegin);
            }
        }
        HSSFRow newrow = sheet.createRow(num + i);
//        newrow.setHeight(templateRow.getHeight());
        if(map.size() > 0){
            //加入序号列
            map.put("_index", i+1);
        }
        for (int j = 0; j < templateRow.getLastCellNum(); j++) {
            HSSFCell cell = templateRow.getCell(j);
            if(cell == null){
                continue;
            }
            HSSFCell createCell = newrow.createCell(j);
            createCell.setCellStyle(cell.getCellStyle());
            String stringCellValue = cell.toString();
            if (isData(stringCellValue)) {
                HSSFConditionalFormatting cellConditionFormating = getCellConditionFormating(
                        sheet, cell);
                // TODO:每次都要算，待优化
                if (formatCondition && null != cellConditionFormating) {
                    CellRangeAddress[] range = { new CellRangeAddress(
                            createCell.getRowIndex(), createCell.getRowIndex(),
                            createCell.getColumnIndex(), createCell
                                    .getColumnIndex()) };
                    int rn = cellConditionFormating.getNumberOfRules();
                    HSSFConditionalFormattingRule[] rule = new HSSFConditionalFormattingRule[rn];
                    for (int l = 0; l < rule.length; l++) {
                        rule[l] = cellConditionFormating.getRule(l);
                    }
                    sheet.getSheetConditionalFormatting()
                            .addConditionalFormatting(range, rule);
                }
                Object value = null;
                if(map.size() > 0){
                    value = getJsValue(stringCellValue, map);
                }
                if (value instanceof Date) {
                    value = new SimpleDateFormat("yyyy-MM-dd")
                            .format((Date) value);
                }
                if (value != null && isNumber(value)) {
                    createCell.setCellValue(Double
                            .parseDouble(value.toString()));
                } else {
                    createCell.setCellValue(new HSSFRichTextString(changeToString(value)));
                }

            }
        }
        return newrow;
    }

    private HSSFRow getLastMergeRow(HSSFRow currentRow, HSSFCell templateCell, int max){
        if(currentRow.getRowNum() == max){
            return currentRow;
        }
        HSSFSheet sheet = templateCell.getSheet();
        HSSFRow row = sheet.getRow(currentRow.getRowNum() + 1);
        HSSFCell currentCell = currentRow.getCell(templateCell.getColumnIndex());
        HSSFCell cell = row.getCell(templateCell.getColumnIndex());
        if(cell != null && cell.toString().equals(currentCell.toString())){
            return getLastMergeRow(row, templateCell, max);
        }else{
            return currentRow;
        }
    }

    private void mergeCellIfNeed(HSSFRow templateRow, int endRow){

        for (int j = 0; j < templateRow.getLastCellNum(); j++) {
            HSSFCell templateCell = templateRow.getCell(j);
            if(templateCell == null){
                continue;
            }
            HSSFComment comment = templateCell.getCellComment();
            if(endRow >= templateCell.getRowIndex() && comment != null){
                String text =  comment.getString().toString();
                if(MERGE_TAG.equals(text.trim())){
                    HSSFSheet sheet = templateCell.getSheet();
                    HSSFRow firstRow = templateRow;
                    while(true){
                        HSSFRow lastRow = getLastMergeRow(firstRow, templateCell, endRow);
                        if(lastRow.getRowNum() > firstRow.getRowNum()){
                            CellRangeAddress region =  new CellRangeAddress(firstRow.getRowNum(),lastRow.getRowNum(),
                                    templateCell.getColumnIndex(), templateCell.getColumnIndex());
                            sheet.addMergedRegion(region);
                        }
                        if(lastRow.getRowNum() >= endRow){
                            break;
                        }
                        firstRow = sheet.getRow(lastRow.getRowNum() + 1);
                    }
                }
            }
            templateCell.setCellComment(null);
        }
    }

    private HSSFConditionalFormatting getCellConditionFormating(
            HSSFSheet sheet, HSSFCell cell) {
        int n = sheet.getSheetConditionalFormatting()
                .getNumConditionalFormattings();
        for (int i = 0; i < n; i++) {
            HSSFConditionalFormatting condition = sheet
                    .getSheetConditionalFormatting()
                    .getConditionalFormattingAt(i);
            for (CellRangeAddress range : condition.getFormattingRanges()) {
                if (cell.getRowIndex() >= range.getFirstRow()
                        && cell.getRowIndex() <= range.getLastRow()
                        && cell.getColumnIndex() >= range.getFirstColumn()
                        && cell.getColumnIndex() <= range.getLastColumn()) {
                    return condition;
                }
            }
        }
        return null;
    }

    private boolean isNumber(Object data) {

        return data instanceof BigDecimal || data instanceof Integer
                || data instanceof Double || data instanceof Float;
    }

    private String changeToString(Object data) {
        if (data == null) {
            return "";
        }
        return data.toString();
    }

    private static String getSourceName(String value){
        String regx = SOURCENAME_REG;
        Pattern p = Pattern.compile(regx);
        Matcher matcher = p.matcher(value);
        if(matcher.find()){
            return matcher.group(1);
        }
        return null;
       }

    private static boolean isData(String value) {
        String regx = "\\$\\{.+\\}";
        Pattern p = Pattern.compile(regx);
        Matcher matcher = p.matcher(value);
        return matcher.find();
    }

    // 判断是否是field
    private static boolean isField(String value) {
        if(value == null || value.length() == 0){
            return false;
        }
        String regx = "\\$f\\{.+\\}";
        Pattern p = Pattern.compile(regx);
        Matcher matcher = p.matcher(value);
        return matcher.find();
    }

    private Object getJsValue(String exp, Map<String, Object> data) throws ScriptException {
        StringTemplate tem = new StringTemplate(exp,
                StringTemplate.REGEX_PATTERN_JAVA_STYLE);
        for (String name : tem.getVariables()) {
            // 纯数字的变量特殊处理
            if (name.matches("\\d+")) {
                tem.setVariable(name, "f" + name);
                engin.put("f" + name, data.get(name));
            } else {
                tem.setVariable(name, name);
                engin.put(name, data.get(name));
            }
        }
        return engin.eval(tem.getParseResult());

    }

    private Object getFieldJsValue(String exp, Map<String, Object> data) throws ScriptException {
        StringTemplate tem = new StringTemplate(exp,
                StringTemplate.REGEX_PATTERN_JAVA_FIELD_STYLE);
        for (String name : tem.getVariables()) {
            // 纯数字的变量特殊处理
            Object value = data.get(name);
            if(null == value){
                value = "";
            }
            if (name.matches("\\d+")) {
                tem.setVariable(name, "f" + name);
                engin.put("f" + name, value);
            } else {
                tem.setVariable(name, name);
                engin.put(name, value);
            }
        }
        return engin.eval(tem.getParseResult());

    }

    private static void updateFormula(HSSFWorkbook wb){
    	int sheetNum = wb.getNumberOfSheets();
        for(int k = 0; k < sheetNum ; k ++){
           HSSFSheet sheet = wb.getSheetAt(k);
           int rownum = sheet.getLastRowNum();
           for (int i = 0; i < rownum + 1; i++) {
               HSSFRow row = sheet.getRow(i);
               if(row == null){
                   continue;
               }
               int cells = row.getLastCellNum();
               for (int j = 0; j < cells; j++) {
                   HSSFCell cell = row.getCell(j);
                   if (cell == null) {
                       continue;
                   }
                   if(cell.getCellType()==HSSFCell.CELL_TYPE_FORMULA){
                	   wb.getCreationHelper().createFormulaEvaluator().evaluateFormulaCell(cell);
                   }
               }
           }
        }
    }

}
