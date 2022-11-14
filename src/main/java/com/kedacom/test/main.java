package com.kedacom.test;

import com.kedacom.util.DocUtilv1;
import com.kedacom.util.DocUtilv2;
import com.kedacom.util.DocUtilv3;
import lombok.SneakyThrows;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @author : lzp
 * @version 1.0
 * @date : 2022/10/31 13:37
 * @apiNote : TODO
 */
public class main {
    @SneakyThrows
    public static void main(String[] args) {
        String path="C:\\Users\\lzp\\Desktop\\doc测试\\所有元素测试.docx";
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
        map.put("JDLX","");
        map.put("CCZP","C:\\Users\\lzp\\Pictures\\pm.jpg");
        map.put("ITEM",2);
        map.put("SZNF",2022);
        map.put("SZYF",11);
        map.put("SZRQ",8);

        ////表格参数
        //List<List<String>> listList=new ArrayList<>();
        //List<String> stringList=new ArrayList<>();
        //stringList.add("1");
        //stringList.add("2");
        //stringList.add("3");
        //stringList.add("4");
        //stringList.add("5");
        //listList.add(stringList);
        //List<String> stringList1=new ArrayList<>();
        //stringList1.add("第二行第一个");
        //stringList1.add("第二个");
        //stringList1.add("第三个");
        //stringList1.add("第四个");
        //stringList1.add("第五个");
        //listList.add(stringList1);
        //map.put("JSZJQD",listList);

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
