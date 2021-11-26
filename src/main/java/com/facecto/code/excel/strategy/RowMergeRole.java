package com.facecto.code.excel.strategy;

import lombok.Data;
import lombok.experimental.Accessors;

import java.util.List;

/**
 * @author Jon So, https://cto.pub, https://github.com/facecto
 * @version v1.0.0 (2021/11/25)
 */
@Data
@Accessors(chain = true)
public class RowMergeRole {
    /**
     * Skip the merged row. For example：3 ,means skip 3 rows and start from the fourth row.
     */
    private Integer rowSkip;
    /**
     * Total number of rows, Used to check the parameter rowMerges.
     */
    private Integer rowNum;
    /**
     * List of row merge rules
     */
    private List<RowMergeColumnRole> list;
}