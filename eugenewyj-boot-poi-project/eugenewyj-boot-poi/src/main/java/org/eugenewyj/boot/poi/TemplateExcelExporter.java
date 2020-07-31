package org.eugenewyj.boot.poi;

import org.jxls.common.Context;
import org.jxls.util.JxlsHelper;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.OutputStream;

/**
 * TemplateExcelExporter
 *
 * @author eugene
 * @date 2020/7/31
 */
public class TemplateExcelExporter {
    private final File templateFile;
    private final Context context = new Context();

    /**
     * 构造函数
     * @param templateFile
     */
    public TemplateExcelExporter(File templateFile) {
        this.templateFile = templateFile;
        if (!templateFile.exists()) {
            throw new IllegalArgumentException(String.format("模板文件 %s 不存在", templateFile.getPath()));
        }
    }

    /**
     * 构造函数
     * @param templateFileClassPath
     */
    public TemplateExcelExporter(String templateFileClassPath) {
        this.templateFile = new File(this.getClass().getResource(templateFileClassPath).getFile());
        if (!templateFile.exists()) {
            throw new IllegalArgumentException(String.format("模板文件 %s 不存在", templateFileClassPath));
        }
    }

    /**
     * 添加要导出的数据。
     * @param key
     * @param data
     * @return
     */
    public TemplateExcelExporter putData(String key, Object data) {
        context.putVar(key, data);
        return this;
    }

    /**
     * 根据模板导出数据。
     * @param os
     * @throws IOException
     */
    public void export(OutputStream os) throws IOException {
        try (FileInputStream is = new FileInputStream(templateFile)) {
            JxlsHelper.getInstance().processTemplate(is, os, context);
        }
    }
}
