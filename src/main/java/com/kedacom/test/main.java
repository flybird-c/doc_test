package com.kedacom.test;

import com.kedacom.util.DocUtilv1;
import com.kedacom.util.DocUtilv2;
import com.kedacom.util.DocUtilv3;
import lombok.SneakyThrows;

import java.io.File;
import java.io.IOException;
import java.util.*;

/**
 * @author : lzp
 * @version 1.0
 * @date : 2022/10/31 13:37
 * @apiNote : TODO
 */
public class main {
    @SneakyThrows
    public static void main(String[] args) {
        String path="C:\\Users\\lzp\\Desktop\\doc测试\\南方医科大学补充补正协议.docx";
        //List<String> codeList=new ArrayList<>();
        //codeList.add("JSZJQD");
        //List codeList_v1 = DocUtilv1.getDocCodes(codeList, path);
        //System.out.println(codeList_v1);
        //List codeList_v2 = DocUtilv2.getDocCodes(codeList, path);
        //System.out.println(codeList_v2);
        //List codeList_v3 = DocUtilv3.getDocCodes(codeList, path);
        //System.out.println(codeList_v3);

        Map<String,Object> map=new HashMap<>();
        map.put("JDYT2","刑事,行政");
        map.put("BQYY","委托方要求更改/补充委托鉴定事项");
        map.put("JDLX","");
        map.put("CCZP","C:\\Users\\lzp\\Pictures\\pm.jpg");
        map.put("ITEM",2);
        map.put("SZNF",2022);
        map.put("SZYF",11);
        map.put("SZRQ",8);

        //表格参数
        //表格参数
        //表格参数
        //表格参数
        //表格参数
        List<List<Object>> listList=new ArrayList<>();
        listList.add(Arrays.asList("","",1,2,3,4,5,6,9,10,11,12));
        listList.add(Arrays.asList("","",12,2,3,4,5,6,9,10,11,12));
        listList.add(Arrays.asList("","",13,2,3,4,5,6,9,10,11,12));
        listList.add(Arrays.asList("","",14,2,3,4,5,6,9,10,11,12));
        listList.add(Arrays.asList("","",15,2,3,4,5,6,9,10,11,12));
        listList.add(Arrays.asList("","",16,2,3,4,5,6,9,10,11,12));
        listList.add(Arrays.asList(21,2,3,4,5));
        listList.add(Arrays.asList(31,2,3,4,5));
        listList.add(Arrays.asList(41,2,3,4,5));
        listList.add(Arrays.asList(51,2,3,4,5));
        listList.add(Arrays.asList(61,2,3,4,5));
        listList.add(Arrays.asList(71,2,3,4,5));
        listList.add(Arrays.asList(81,2,3,4,5));
        listList.add(Arrays.asList(91,2,3,4,5));
        listList.add(Arrays.asList(101,2,3.14f,4,5));
        listList.add(Arrays.asList(111,2,3,4.12,5));
        listList.add(Arrays.asList(121,2,3,4,5L));
        listList.add(Arrays.asList(131,2,3,4,5));
        listList.add(Arrays.asList(141,2,"这是字符串",4,5));
        listList.add(Arrays.asList(151,2,3,4,5));
        listList.add(Arrays.asList(161,2,3,4,null));
        map.put("HDJL",listList);

        //map.put("SZXS","");
        //map.put("XM","");
        //map.put("JSXM","");
        //map.put("JSDH","");
        //String s = replaceWordCode_v1(map, path);
        //System.out.println(s);
        ////String s2 = replaceWordCode_v2(map, path);
        ////System.out.println(s2);
        String s3 = DocUtilv3.replaceWordCode(map, path);
        System.out.println(s3);
    }
    public static  List getCodeList_v1(List<String> codeList,String path){
        return DocUtilv1.getDocCodes(codeList,path);
    }
    public static List getCodeList_v2(List<String> codeList,String path){
        return DocUtilv2.getDocCodes(codeList,path);
    }
    public static String replaceWordCode_v1(Map<String,Object> map, String srcWordPath) throws IOException {
         return DocUtilv1.replaceWordCode(map,srcWordPath);
    }
    public static String replaceWordCode_v2(Map<String,Object> map, String srcWordPath){
        return DocUtilv2.replaceWordCode(map,srcWordPath);
    }

}
