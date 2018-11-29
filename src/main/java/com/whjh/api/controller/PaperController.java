package com.whjh.api.controller;

import com.alibaba.fastjson.JSONObject;
import com.whjh.api.service.PaperService;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.annotation.Resource;
import javax.servlet.http.HttpServletResponse;
import java.io.*;

/**
 * @author Lee
 * @date 2018/11/15
 */
@RestController
@RequestMapping("/paper")
public class PaperController {
    @Resource
    private PaperService paperService;

    @PostMapping("/download")
    public void downloadDocx(@RequestBody JSONObject jsonParam, HttpServletResponse response) {
        try {
            String filename = this.paperService.createPaperDocx(jsonParam);
            File file=new File(filename);

            String fileName = file.getName();// 获取日志文件名称
            InputStream fis = new BufferedInputStream(new FileInputStream(file));
            System.out.println(fis.available());
            byte[] buffer = new byte[fis.available()];
            fis.read(buffer);
            fis.close();
            response.reset();
            // 先去掉文件名称中的空格,然后转换编码格式为utf-8,保证不出现乱码,这个文件名称用于浏览器的下载框中自动显示的文件名
            response.addHeader("Content-Disposition", "attachment;filename=" + new String(fileName.replaceAll(" ", "").getBytes("utf-8"),"iso8859-1"));
            response.addHeader("Content-Length", "" + file.length());
            OutputStream os = new BufferedOutputStream(response.getOutputStream());
            response.setContentType("application/octet-stream");
            os.write(buffer);// 输出文件
            os.flush();
            os.close();

            file.delete();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
