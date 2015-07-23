package service;

import org.apache.poi.ss.usermodel.Workbook;

import java.util.List;

/**
 * Created by dhruvr on 21/7/15.
 */
public interface ExcelService {

    /**
     * this method creates a new Excel sheet
     *
     * @param listObjs the list of some pojo that we need to write to excel
     * @param <T>
     * @return the workbook with a created sheet in it
     */
    <T> Workbook create(List<T> listObjs)  throws NoSuchFieldException, IllegalAccessException;

}
