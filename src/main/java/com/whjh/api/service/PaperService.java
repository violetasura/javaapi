package com.whjh.api.service;

import com.alibaba.fastjson.JSONObject;

/**
 * @author Lee
 * @date 2018/11/14
 */
public interface PaperService {
    /**
     * 生成试卷docx
     * @param jsonObject
     * @return 文件名
     */
    String createPaperDocx(JSONObject jsonObject);
}
