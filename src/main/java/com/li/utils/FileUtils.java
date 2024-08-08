package com.li.utils;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.URL;
import java.net.URLConnection;

public class FileUtils extends org.apache.commons.io.FileUtils {

    public static InputStream getInputStream(String fileUrl) {
        InputStream in;
        try {
            if (fileUrl.startsWith("http")) {
                // 网络地址
                URL urlObj = new URL(fileUrl);
                URLConnection urlConnection = urlObj.openConnection();
                urlConnection.setConnectTimeout(30 * 1000);
                urlConnection.setReadTimeout(60 * 1000);
                urlConnection.setDoInput(true);
                in = urlConnection.getInputStream();
            }
            else {
                in = new FileInputStream(fileUrl);
            }
            return in;
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }
}
