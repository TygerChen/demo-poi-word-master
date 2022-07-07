package com.demo.poiword;

import java.io.InputStream;
import java.net.HttpURLConnection;
import java.net.URL;

public class HttpUtils {

    /**
    * @Description 通过图片URL地址得到文件流
    * @param url 图片地址
    * @retuen java.io.InputStream
    * @author Tyger Chen
    * @since 2022/7/7 13:59
    */
    public static InputStream getFileStream(String url){
        try {
            URL httpUrl = new URL(url);
            HttpURLConnection conn = (HttpURLConnection)httpUrl.openConnection();
            conn.setRequestMethod("GET");
            conn.setConnectTimeout(5 * 1000);
            InputStream inStream = conn.getInputStream();//通过输入流获取图片数据
            return inStream;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }

}
