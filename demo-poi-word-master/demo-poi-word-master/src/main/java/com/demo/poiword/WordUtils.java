
package com.demo.poiword;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.*;
import org.springframework.stereotype.Component;

import javax.imageio.ImageIO;
import javax.servlet.http.HttpServletResponse;
import java.awt.image.BufferedImage;
import java.io.*;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

//使用poi导出word模板的操作都封装在这了
@Component
public class WordUtils {
    /**
     * description:导出word模板
     * @param path:word模板路径
     * @param params:模板中需要替换的参数可多个传递 比如 若想文字能够多行，在参数Map<String,Object>中的Object访如List<string>
     *                               若是传递图片 参数Map<String,Object>中的Object为Map<String,Obejct>   若是一个参数处要传入多张图片Object就为List<Map<String,Object>
     *                               在传入多张图片的时候 Map中设置style ,style=1 代表图片并排插入，style=2代表图片竖排插入，设置imgpath，即图片路径
     * @param filename:导出的word文件名
     * @param response:
     * @return void
     * @Author any
     * @Date 2020/8/10 15:29
     */
    public void exportWord(String path, Map<String, Object> params, String filename, HttpServletResponse response) throws IOException, InvalidFormatException {

        InputStream is = new FileInputStream(path);
        //代表一个docx文档
        XWPFDocument doc = new XWPFDocument(is);
        //开始遍历文档中的表格
        iteraTable(params, doc);
        OutputStream os = response.getOutputStream();
        //设置导出的内容是doc
        response.setContentType("application/octet-stream; charset=utf-8");
        response.setHeader("Content-disposition", "attachment; filename=" + filename);
        doc.write(os);
        close(os);
    }

    //开始遍历文档中的表格
    //一个文档包含多个表格，一个表格包含多行，一行包含多列(格)，每一格的内容相当于一个完整的文档
    private void iteraTable(Map<String, Object> params, XWPFDocument doc) throws IOException, InvalidFormatException {
        //XWPFTableRow 代表一个表格的行
        List<XWPFTableRow> rows = null;
        //XWPFTableCell 代表表格的一个单元格
        List<XWPFTableCell> cells = null;
        //获取此文档中的所有表格
        List<XWPFTable> tables = doc.getTables();
        //遍历这个word文档的所有table表格
        for (XWPFTable table : tables) {
            //得到某个表格的所有行
            rows = table.getRows();
            //遍历这个表格的所有行
            for (XWPFTableRow row : rows) {
                //得到某一行的所有单元格
                cells = row.getTableCells();
                //遍历这一行的单元格
                for (XWPFTableCell cell : cells) {
                    //判断该单元格的内容是否是字符串字段
                    if (strMatcher(cell.getText()).find()) {
                        //替换字符串 字符串可以多行 也可以一行
                        //设置单元格的颜色
                        cell.setColor("BED4F1");

                        replaceInStr(cell, params, doc);
                        continue;
                    }
                    //判断该单元格内容是否是需要替换的图片
                    if (imgMatcher(cell.getText()).find()) {
                        //把模板中的内容替换成图片 图片可以多长
                        replaceInImg(cell, params, doc);
                        continue;
                    }
                }
            }
        }
    }

    //返回模板中图片字符串的匹配Matcher类
    private Matcher imgMatcher(String imgstr) {
        Pattern pattern = Pattern.compile("@\\{(.+?)\\}");
        Matcher matcher = pattern.matcher(imgstr);
        return matcher;
    }

    //返回模板中变量的匹配Matcher类
    private Matcher strMatcher(String str) {
        Pattern pattern = Pattern.compile("\\$\\{(.+?)\\}");
        Matcher matcher = pattern.matcher(str);
        return matcher;
    }


    //替换模板Table中相应字段为对应的字符串值
    private void replaceInStr(XWPFTableCell cell, Map<String, Object> params, XWPFDocument doc) {
        //获取单元格中变量的名字---去掉${ }
        String key = cell.getText().substring(2, cell.getText().length() - 1);
        //两种数据类型  一种是直接String  还有一种是List<String>
        //获取map中存储的值的类型：null/string/list<string>
        Integer datatype = getMapStrDataTypeValue(params, key);
        //获取一个单元格里的所有内容(所有段落)
        List<XWPFParagraph> parags = cell.getParagraphs();
        //先清空单元格中所有的段落
        for (int i = 0; i < parags.size(); i++) {
            cell.removeParagraph(i);
        }
        //0 代表为空，无值
        if (datatype.equals(0)) {
            return;
        } else if (datatype.equals(2)) {
            //如果类型是2 说明数据类型是List<String>
            List<String> strs = (List<String>) params.get(key);
            Iterator<String> iterator = strs.iterator();
            while (iterator.hasNext()) {
                //给单元格加一个段落对象
                XWPFParagraph para = cell.addParagraph();
                //在段落里加入文字对象，文字对象可以设置样式
                XWPFRun run = para.createRun();
                //设置文字对象的文本内容
                run.setText(iterator.next());
                //加粗
                run.setBold(true);
                //字体大小
                run.setFontSize(20);
                //颜色
                run.setColor("f40001");
                run.setStrikeThrough(true);
            }
            //如果类型是3 说明数据类型是String
        } else if (datatype.equals(3)) {
            String str = params.get(key).toString();
            //设置文字对象的文本内容
            cell.setText(str);
        }
    }

    //替换模板Table中相应字段为图片
    private void replaceInImg(XWPFTableCell cell, Map<String, Object> params, XWPFDocument doc) throws IOException, InvalidFormatException {
        //拿参数 判断一下是什么类型 怎么处理   一般来说Map中只有两种类型 一种是一个模板图片变量放一张图片 只有一个Map  一种是一个模板变量中放多涨图片List<Map> 还有一种是不存在的情况
        String key = cell.getText().substring(2, cell.getText().length() - 1);
        //获取mapimg中存储的值的类型：null/Map/list<Map>
        Integer datatype = getMapImgDataType(params, key);
        //获取一个单元格中所有的段落
        List<XWPFParagraph> parags = cell.getParagraphs();
        //先清空单元格中所有的段落
        for (int i = 0; i < parags.size(); i++) {
            cell.removeParagraph(i);
        }
        //0代表为null，无值，不作处理
        if (datatype.equals(0)) {
            return;
        //1代表为map类型 单张图片
        } else if (datatype.equals(1)) {
            //处理单张图片
            //给单元格添加一个段落
            XWPFParagraph parag = cell.addParagraph();
            Map<String, Object> pic = (Map<String, Object>) params.get(key);
            //在添加的段落中插入图片
            insertImg(pic, parag);
        } else if (datatype.equals(2)) {
            //处理多张图片
            List<Map<String, Object>> pics = (List<Map<String, Object>>) params.get(key);
            Iterator<Map<String, Object>> iterator = pics.iterator();
            XWPFParagraph para = cell.addParagraph();
            Integer count = 0;
            while (iterator.hasNext()) {
                Map<String, Object> pic = iterator.next();
                //图片并排插入
                //if (Integer.parseInt(pic.get("style").toString()) == 1) {
                //    //不做处理 这段代码可注释
                //}
                //图片竖排插入
                //一排代表一个段落，竖排插入多个需要添加多个段落
                if (Integer.parseInt(pic.get("style").toString()) == 2) {
                    if (count > 0) {
                        para = cell.addParagraph();
                    }
                }
                insertImg(pic, para);
                count++;
            }
        }
    }


    //插入图片   run创建
    private void insertImg(Map<String, Object> pic, XWPFParagraph para) throws IOException, InvalidFormatException {
        if (pic.get("imgpath").toString() == null)
            return;
        //获取图片路径
        String picpath = pic.get("imgpath").toString();
        //字节输入流
        InputStream is = null;
        //可处理图片的属性
        BufferedImage bi = null;
        //代表是服务器中的图片
        if (picpath.startsWith("http")) {
            is = HttpUtils.getFileStream(picpath);
            bi = ImageIO.read(HttpUtils.getFileStream(picpath));
        }
        //代表的是本地库中的图片
        else {
            is = new FileInputStream(picpath);
            bi = ImageIO.read(new File(picpath));
        }
        //创建一个文本对象
        XWPFRun run = para.createRun();
        //原图片的长宽
        Integer width = bi.getWidth();
        Integer heigh = bi.getHeight();
        Double much = 80.0 / width;
        //图片按宽80 比例缩放
        //向文本对象中添加图片
        run.addPicture(is, getPictureType(picpath.substring(picpath.lastIndexOf(".") + 1)), "", Units.toEMU(80), Units.toEMU(heigh * much));
        //图片原长宽
//        run.addPicture(is,getPictureType(pic.get("picType").toString()),"",Units.toEMU(width),Units.toEMU(heigh));
        close(is);
        bi = null;
    }

    //插入图片  自定义 doc方法
//    private static void insertImg(Map<String,Object> pic,XWPFParagraph para,CustomXWPFDocument doc) throws FileNotFoundException, InvalidFormatException {
//        InputStream is=null;
//        if(pic.get("imgpath").toString()==null)
//            return;
//        is=new FileInputStream(pic.get("imgpath").toString());
//        byte[] bytes=inputStream2ByteArray(is,true);
//        doc.addPictureData(bytes, getPictureType(pic.get("picType").toString()));
//        doc.createPicture(doc.getAllPictures().size() - 1, 30, 360,para);
//        close(is);
//    }


    //处理模板中的变量名字，去掉${}  然后根据这个变量名在参数map中查找对应的Value值
    private Integer getMapStrDataTypeValue(Map<String, Object> params, String key) {
        if (params.get(key) == null) {
            return 0;
        } else if (params.get(key) instanceof List) {
            return 2;
        } else if (params.get(key) instanceof String) {
            return 3;
        } else {
            throw new RuntimeException("Str data type error!");
        }

    }

    //从params参数中找到模板中key对应的value值
    private Integer getMapImgDataType(Map<String, Object> params, String key) {
        if (params.get(key) == null) {
            return 0;
        } else if (params.get(key) instanceof Map) {
            return 1;
        } else if (params.get(key) instanceof List) {
            return 2;
        } else {
            throw new RuntimeException("image data type error!");
        }
    }


    /**
     * 根据图片类型，取得对应的图片类型代码
     *
     * @param picType
     * @return int
     */
    private int getPictureType(String picType) {
        int res = XWPFDocument.PICTURE_TYPE_PICT;
        if (picType != null) {
            if (picType.equalsIgnoreCase("png")) {
                res = XWPFDocument.PICTURE_TYPE_PNG;
            } else if (picType.equalsIgnoreCase("dib")) {
                res = XWPFDocument.PICTURE_TYPE_DIB;
            } else if (picType.equalsIgnoreCase("emf")) {
                res = XWPFDocument.PICTURE_TYPE_EMF;
            } else if (picType.equalsIgnoreCase("jpg") || picType.equalsIgnoreCase("jpeg")) {
                res = XWPFDocument.PICTURE_TYPE_JPEG;
            } else if (picType.equalsIgnoreCase("wmf")) {
                res = XWPFDocument.PICTURE_TYPE_WMF;
            }
        }
        return res;
    }

    /**
     * 将输入流中的数据写入字节数组
     *
     * @param in
     * @return
     */
    public byte[] inputStream2ByteArray(InputStream in, boolean isClose) {
        byte[] byteArray = null;
        try {
            //获得输入流中还有多少字节可读取
            int total = in.available();
            byteArray = new byte[total];
            in.read(byteArray);
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            if (isClose) {
                try {
                    in.close();
                } catch (Exception e2) {
                    e2.getStackTrace();
                }
            }
        }
        return byteArray;
    }


    /**
     * 关闭输入流
     *
     * @param is
     */
    private void close(InputStream is) {
        if (is != null) {
            try {
                is.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    /**
     * 关闭输出流
     *
     * @param os
     */
    private void close(OutputStream os) {
        if (os != null) {
            try {
                os.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }
}
