
package com.lixi.util.excel;


import java.io.IOException;
import java.io.InputStream;
import java.text.DecimalFormat;
import java.util.Iterator;
import java.util.List;
import java.util.regex.Pattern;


public class StringUtil {

    /**
     * 判断是否为空字符串（null或者空串）
     */
    public static boolean isEmpty(Object obj) {
        return obj == null || "".equals(obj);
    }

    /**
     * 判断是否为非空字符串（不为null或者空串）
     */
    public static boolean isNotEmpty(Object obj) {
        return !isEmpty(obj);
    }

    /**
     * 替换字符串
     */
    public static String replaceAll(String strSource, String strFrom, String strTo) {
        String strDest = "";
        int intFromLen = strFrom.length();
        int intPos;
        if (strSource == null) {
            strDest = "";
            strSource = "";
        } else {
            while ((intPos = strSource.indexOf(strFrom)) != -1) {
                strDest = strDest + strSource.substring(0, intPos);
                strDest = strDest + strTo;
                strSource = strSource.substring(intPos + intFromLen);
            }
        }
        strDest = strDest + strSource;
        return strDest;
    }

    /**
     * 替换第一匹配字符
     */
    public static String replaceFirst(String src, String match, String as) {
        String result = replaceNull(src);

        int start = result.indexOf(replaceNull(match));
        if (start != -1 && StringUtil.isNotEmpty(match)) {
            result = new StringBuffer(result).replace(start, start + match.length(), replaceNull(as)).toString();
        }
        return result;
    }

    /**
     * 替换最后一个匹配字符
     */
    public static String replaceLast(String src, String match, String as) {
        String result = replaceNull(src);

        result = reverseStr(result);
        match = reverseStr(match);
        as = reverseStr(as);

        result = reverseStr(replaceFirst(result, match, as));

        return result;
    }

    /**
     * 字符串反转
     */
    public static String reverseStr(String source) {
        if (isEmpty(source)) {
            return source;
        } else {
            return (new StringBuffer(source).reverse()).toString();
        }
    }


    /**
     * 根据分隔符合并List到字符串
     */
    public static String joinString(List<String> lstInput, String strToken) {
        if (lstInput == null || lstInput.size() == 0) {
            return "";
        }
        StringBuffer sb = new StringBuffer();
        for (Iterator<String> iterator = lstInput.iterator(); iterator.hasNext(); ) {
            sb.append(iterator.next().toString());
            if (iterator.hasNext()) {
                sb.append(strToken);
            }
        }
        return sb.toString();
    }

    /**
     * 重构字符串，如果为null，返回空串
     */
    public static String replaceNull(Object obj) {
        if (obj == null) {
            return "";
        } else {
            return obj.toString();
        }
    }

    /**
     * inputStream to  String
     * **/
    public static String inputStreamToString(InputStream in){
    	return inputStreamToString(in, null);
    }
    
    /**
     * inputStream to  String
     * **/
    public static String inputStreamToString(InputStream in, String charset){
    	
    	StringBuffer out =new  StringBuffer(); 
        byte[]  b = new  byte[4096]; 
        try {
			for(int n;(n=in.read(b))!=-1;){ 
			   if(null == charset){
				   out.append(new String(b,0,n)); 
			   }else{
				   out.append(new String(b,0,n,charset)); 
			   }
			}
		}catch(IOException e){
			e.printStackTrace();
		} 
        return  out.toString(); 
    }
    
    /**
     * 格式化 整数以一定的位数输出
     * length:格式化输出的整数的位数
     * number：要格式化的数字
     * **/
    public static String formatInteger(int length,int number){
    	
    	String reg="";
    	for(int i=0;i<length;i++){
    		reg+="0";
    	}
    	return new DecimalFormat(reg).format(number);
    }
    
    
    public static boolean isNumeric(String value){
    	if(isEmpty(value)){
    		return false;
    	}
    	Pattern pattern = Pattern.compile("^(-?\\d+)(\\.\\d+)?$");
        return pattern.matcher(value).matches();
    }
    
    /**
     * 判断一个字符串是否包含
     * @param value
     * @return
     */
    public static boolean containsLetter(String value){
        return Pattern.compile("(?i)[a-z]").matcher(value).find();
    }
    
}
