package com.dfei.tools.tools.exportUtils.excelUtils.utils;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

/**
 * @author dfei
 * @ClassName Util.java
 * @createTime 2022/1/18 14:25
 * @Description TODO
 */
public class Util {

    public static List<String> getMapKeys(Map<String, Object> map) {
        List<String> keys = new ArrayList<>();
        map.keySet().stream().forEach(e -> keys.add(e));
        return keys;
    }

    public static void main(String[] args) {

    }

}
