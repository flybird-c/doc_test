package com.kedacom.util;

import org.junit.jupiter.api.Test;

import java.util.HashMap;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.*;

/**
 * @author : lzp
 * @version 1.0
 * @date : 2022/11/1 11:11
 * @apiNote : TODO
 */
class DocUtilv2Test {
    @Test
    public void testItem() {
        Map<String,Object> map=new HashMap<>();
        map.put("ITEM",1);
        map.put("JDYT2","行政");
        Map<String, Object> stringObjectMap = DocUtilv2.handItemCodes(map);
        System.out.println(stringObjectMap);
    }
}