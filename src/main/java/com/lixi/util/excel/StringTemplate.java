package com.lixi.util.excel;

import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class StringTemplate {


    /**Java 风格的模板变量引用字符串: ${...}*/
    public static final String REGEX_PATTERN_JAVA_STYLE = "\\$\\{([^\\}]*)\\}";
    
    public static final String REGEX_PATTERN_JAVA_FIELD_STYLE = "\\$f\\{([^\\}]*)\\}";
    
    /**Unix shell 风格的模板变量引用字符串: $...*/
    public static final String REGEX_PATTERN_UNIX_SHELL_STYLE = "\\$([a-zA-Z_]+)";
    
    private Map varExpMap = new HashMap();      //Map <varName:String, varExpression:String>
    private Map varValueMap = new HashMap();    //Map <varName:String, varValue:String>
    private String templateString;

    @SuppressWarnings("unchecked")
    public StringTemplate(String templateString, String templateRegex){
        Pattern p = Pattern.compile(templateRegex);
        Matcher m = p.matcher(templateString);
        while(m.find()){
            if (m.groupCount()==1){
                //group[0]: the all match string, group[1]: the group
                varExpMap.put(m.group(1), m.group(0));
            }else{
                throw new RuntimeException(
                        "Pattern '"+templateRegex+"' is not valid, it must contain and only contain one regex group.");
            }
        }
        this.templateString = templateString;
    }
    
    @SuppressWarnings("unchecked")
    public String[] getVariables(){
        return (String[])varExpMap.keySet().toArray(new String[]{});
    }
    
    @SuppressWarnings("unchecked")
    public void setVariable(String varName, String varValue){
        varValueMap.put(varName, varValue);
    }
    
    public String getParseResult(){
        Iterator itr = varExpMap.entrySet().iterator();
        String tmp = this.templateString;
        while (itr.hasNext()){
            Map.Entry en = (Map.Entry)itr.next();
            String varName = (String)en.getKey();
            String varExp = (String)en.getValue();
            String val = (String)varValueMap.get(varName);
            if (null!=val){
                tmp = replace(tmp, varExp, val);
            }else{
                //在对应变量没有复制的情况下, 字符串保持原样.
            }
        }
        return tmp;
    }
    
    public void reset(){
        this.varValueMap.clear();
    }
    
    public static String replace(String template, String placeholder, String replacement) {

        return replace(template, placeholder, replacement, false);
    }

    public static String replace(String template, String placeholder, String replacement, boolean wholeWords) {

        int loc = template.indexOf(placeholder);

        if (loc < 0) {
            return template;
        }
        final boolean actuallyReplace = !wholeWords || ((loc + placeholder.length()) == template.length())
                || !Character.isJavaIdentifierPart(template.charAt(loc + placeholder.length()));
        String actualReplacement = actuallyReplace ? replacement : placeholder;

        return new StringBuffer(template.substring(0, loc)).append(actualReplacement).append(
                replace(template.substring(loc + placeholder.length()), placeholder, replacement, wholeWords))
                .toString();
    }

}
