package com.facecto.code.excel.strategy;

import com.alibaba.excel.metadata.Head;
import com.alibaba.excel.write.merge.AbstractMergeStrategy;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

/**
 * @author Jon So, https://cto.pub, https://github.com/facecto
 * @version v1.0.0 (2021/11/25)
 */
public class MergeStrategy extends AbstractMergeStrategy {

    private RowMergeRole rowMergeRole;
    private Sheet sheet;
    private Integer skipRow =0;

    public MergeStrategy(RowMergeRole rowMergeRole) {
        this.rowMergeRole =rowMergeRole;
        this.skipRow = rowMergeRole.getRowSkip();
    }

    private void mergeRow(Integer columnIndex){
        Integer rowCurrent = rowMergeRole.getRowSkip();
        for(RowMergeColumnRole role: rowMergeRole.getList()){
            if(columnIndex >= role.getColumnStart() && columnIndex <= role.getColumnEnd()){
                for (Integer count: role.getRowMerges()){
                    if(count>1){
                        CellRangeAddress cellRangeAddress = new CellRangeAddress(rowCurrent,rowCurrent+count-1,columnIndex,columnIndex);
                        sheet.addMergedRegionUnsafe(cellRangeAddress);
                        rowCurrent+=count;
                    }
                    else {
                        rowCurrent+=count;
                    }
                }
            }
        }
    }

    @Override
    protected void merge(Sheet sheet, Cell cell, Head head, Integer integer) {
        this.sheet = sheet;
        int rowIndex = cell.getRowIndex();
        int columnIndex = cell.getColumnIndex();
        Row row = cell.getRow();
        if (rowIndex == this.skipRow) {
            mergeRow(cell.getColumnIndex());
        }
    }
}