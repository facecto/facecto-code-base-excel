package com.facecto.code.excel.strategy;

import lombok.Data;
import lombok.experimental.Accessors;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.VerticalAlignment;

/**
 * @author Jon So, https://cto.pub, https://github.com/facecto
 * @version v1.0.0 (2021/11/25)
 */
@Data
@Accessors(chain = true)
public class BaseStyle {
    private Short fontSize = 12;
    private String fontName = "宋体";
    private Boolean hasBold = false;
    private IndexedColors foregroundColor = IndexedColors.WHITE;
    private IndexedColors backgroundColor ;
    private HorizontalAlignment horizontalAlignment = HorizontalAlignment.CENTER;
    private VerticalAlignment verticalAlignment = VerticalAlignment.CENTER;
}
