package tools;

import java.util.ResourceBundle;

public class ReadProperty {
    /**
     * 使用静态方法可以直接读取application属性文件中的值
     * ReadProperty.readValue(key)即可返回值
     */
    private static String filename  =  "application";
    public static String  readValue(String key){
        ResourceBundle bundle =  ResourceBundle.getBundle(filename);
        return bundle.getString(key);
    }
}
