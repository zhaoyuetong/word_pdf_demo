package com.zyt.utils.utils;

import com.aspose.words.License;
import com.aspose.words.SaveFormat;
import org.springframework.core.io.ClassPathResource;
import org.springframework.core.io.support.PathMatchingResourcePatternResolver;
import org.springframework.core.io.support.ResourcePatternResolver;

import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;
import org.springframework.core.io.Resource;

public class WordToPdf {
    /**
     *
     * @Title: getLicense
     * @Description:验证license许可凭证
     * @return boolean 返回验证License状态
     */
    private static boolean getLicense() {
        boolean result = true;
        try {
//            Resource resource = new ClassPathResource("./license.xml");//本地

            //本地获取license.xml配置文件
            Resource resource = new ClassPathResource("license.xml");
            InputStream is = resource.getInputStream();
            new License().setLicense(is);
            //服务器获取license.xml配置文件
//            ResourcePatternResolver resolver = new PathMatchingResourcePatternResolver();
//            Resource[] resources =resolver.getResources("file:"+System.getProperty("user.dir")+"/license.xml");
//            InputStream is = resources[0].getInputStream();
//            new License().setLicense(is);
        } catch (Exception e) {
            result = false;
            e.printStackTrace();
        }
        return result;
    }
    /**
     *
     * @Title: wordConverterToPdf
     * @Description: word转pdf(aspose转换)
     * @param wordPath word 全路径，包括文件全称
     * @param pdfPath pdf 保存路径，包括文件全称
     * @return boolean 返回转换状态
     */
    public static File wordConverterToPdf(String wordPath, String pdfPath) {
        System.out.println("===================aspose开始转换=======================");
        //开始时间
        long start = System.currentTimeMillis();
        boolean bool = false;
        // 验证License,若不验证则转化出的pdf文档会有水印产生
        if (!getLicense()) return null;
        try {
            File file = new File(pdfPath);
            if (!file.getParentFile().exists()) {
                file.getParentFile().mkdirs();
            }
            FileOutputStream os = new FileOutputStream(new File(pdfPath));// 新建一个pdf文档输出流
            com.aspose.words.Document doc = new com.aspose.words.Document(wordPath);// Address是将要被转化的word文档
            doc.save(os, SaveFormat.PDF);// 全面支持DOC, DOCX, OOXML, RTF HTML, OpenDocument, PDF, EPUB, XPS, SWF 相互转换
            os.close();
            bool = true;
            System.out.println("(aspose)word转换PDF完成，用时"+(System.currentTimeMillis()-start)/1000.0+"秒");
            System.out.println(file);
            return file;
        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }
}
