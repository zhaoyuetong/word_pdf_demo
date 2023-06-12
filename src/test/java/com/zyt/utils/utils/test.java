package com.zyt.utils.utils;

import org.junit.internal.Classes;
import org.junit.jupiter.api.Test;
import org.junit.runner.RunWith;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.context.ContextConfiguration;
import org.springframework.test.context.junit4.SpringJUnit4ClassRunner;

import java.util.HashMap;
import java.util.Map;

@SpringBootTest(classes = {WordToPdf.class,WordUtils.class})
public class test {
    //world模板
    public static String templatePath = "D:\\workspaces\\test\\template\\template.docx";
    //填充数据后模板地址
    public static String fileDir = "D:\\workspaces\\test\\template\\template.docx";


    public static String DOC_PATH = "D:\\workspaces\\test\\DwOpenAgree.docx";
    public static String PDF_PATH = "D:\\workspaces\\test\\DwOpenAgree.pdf";
    @Test
    public void testWordToPdf(){
        WordToPdf.wordConverterToPdf(DOC_PATH,PDF_PATH);
    }
    @Test
    public void testWord(){
        Map<String, Object> params = new HashMap<>();
        // 渲染文本
        params.put("name", "张三");
        params.put("date", "20220202");
        params.put("part", "开发部");
        params.put("job", "人事");
//        ...
        // 渲染图片
//        params.put("picture", new PictureRenderData(120, 120, "D:\\wx.png"));
        // TODO 渲染其他类型的数据请参考官方文档
        //生成的文件名
        String fileName = "zdd2";
        String wordPath = WordUtils.createWord(templatePath, fileDir, fileName, params);
        System.out.println("生成文档路径：" + wordPath);
    }
}
