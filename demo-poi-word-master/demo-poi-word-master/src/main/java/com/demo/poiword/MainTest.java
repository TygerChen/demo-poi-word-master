package com.demo.poiword;


import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
/**
* @author Tyger Chen
* @since 2022/7/7 13:59
*/
@RestController
public class MainTest {
    @Autowired
    private WordUtils wordUtils;

    @RequestMapping("/test/mydocx")
    public void exportWord(HttpServletResponse response) throws IOException, InvalidFormatException {

        //模板地址
        String path="F:\\testmodel\\model.docx";
        //传入模板参数
        Map<String,Object> params=new HashMap<>();
        //模板的导出文件名字
        String filename="mydocxtest.docx";

        //一个单元格内一行字
        //${name} ${sex}
        params.put("name","张三");
        params.put("sex","男");

        //一个单元格内多行字
        //${hobby}
        List<String> hobby=new ArrayList<>();
        hobby.add("1、打篮球");
        hobby.add("2、打羽毛球");
        hobby.add("3、游泳");
        params.put("hobby",hobby);


        //图片参数1 多张图片 图片多行竖排排列
        //@{workimg}
        List<Map<String,Object>> imgs1List=new ArrayList<>();
        Map<String,Object> img=new HashMap<>();
        img.put("style",2);
        img.put("imgpath","F:\\mytestimg\\testimg1.jpg");
        imgs1List.add(img);
        img=new HashMap<>();
        img.put("style",2);
        //网上图片库。
        img.put("imgpath","https://gimg2.baidu.com/image_search/src=http%3A%2F%2Fc-ssl.duitang.com%2Fuploads%2Fitem%2F202105%2F29%2F20210529001057_aSeLB.thumb.1000_0.jpeg&refer=http%3A%2F%2Fc-ssl.duitang.com&app=2002&size=f9999,10000&q=a80&n=0&g=0n&fmt=auto?sec=1659753892&t=705d99502799699cf40effcd143a3b9b");
        imgs1List.add(img);
        params.put("workimg",imgs1List);

        //图片参数2 多张图片 一行内横排排列
        //${signimg}
        List<Map<String,Object>> imgs2List=new ArrayList<>();
        Map<String,Object> img2=new HashMap<>();
        img2.put("style",1);
        img2.put("imgpath","F:\\mytestimg\\sign1.png");
        imgs2List.add(img2);
        img2=new HashMap<>();
        img2.put("style",1);
        img2.put("imgpath","F:\\mytestimg\\sign2.png");
        imgs2List.add(img2);
        img2=new HashMap<>();
        img2.put("style",1);
        img2.put("imgpath","F:\\mytestimg\\sign3.png");
        imgs2List.add(img2);
        params.put("signimg",imgs2List);


        //图片参数3 单张图片
        //${otherimg}
        Map<String,Object> img3=new HashMap<>();
        img3.put("imgpath","F:\\mytestimg\\testimg1.jpg");
        params.put("otherimg",img3);

        wordUtils.exportWord(path,params,filename,response);
    }



}
