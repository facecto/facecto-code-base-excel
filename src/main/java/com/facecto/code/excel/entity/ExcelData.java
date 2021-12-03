package com.facecto.code.excel.entity;

import lombok.Data;
import lombok.experimental.Accessors;

import java.util.List;

/**
 * @author Jon So, https://cto.pub, https://github.com/facecto
 * @version v1.0.0 (2021/12/03)
 */
@Data
@Accessors(chain = true)
public class ExcelData<X,Y> {
    private List<X> sheetHeadList;
    private List<SheetBody<Y>> sheetBodyList;
    private List<String> sheetNameList;
}
