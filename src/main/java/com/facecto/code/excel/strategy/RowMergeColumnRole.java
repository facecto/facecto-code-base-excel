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
public class RowMergeColumnRole {
    /**
     * The starting column of the merged row. The minimum number is 0, indicating the first column.
     * For example, 1, means it starts in the second column.
     */
    private Integer columnStart;
    /**
     * The end of the merged row. The minimum number is 0 and contains the current row
     * For example, 2, means merge the third column
     */
    private Integer columnEnd;
    /**
     * List of merged numbers.
     * For example, 3 means merge 3 rows.
     */
    private List<Integer> rowMerges;
}