package com.li.utils.word;

import java.io.Serializable;

/**
 * 占位符信息
 * @author longxiong
 * @date 2024/8/7 14:25:03
 */
public class Placeholder implements Serializable {

    /** 占位符名称 */
    private String key;

    /** 占位符内容 */
    private String value;

    /** 占位符内容类型（1-文本 2-图片） */
    private int type;

    public String getKey() {
        return key;
    }

    public void setKey(String key) {
        this.key = key;
    }

    public String getValue() {
        return value;
    }

    public void setValue(String value) {
        this.value = value;
    }

    public int getType() {
        return type;
    }

    public void setType(int type) {
        this.type = type;
    }
}
