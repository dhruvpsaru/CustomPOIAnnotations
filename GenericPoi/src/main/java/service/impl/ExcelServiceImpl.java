package service.impl;

import annotations.ExcelColumn;
import annotations.ExcelReport;
import beans.LabelAndFields;
import com.google.common.base.Preconditions;
import com.sun.istack.internal.NotNull;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import service.ExcelService;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

/**
 * Created by dhruvr on 21/7/15.
 */
public class ExcelServiceImpl implements ExcelService {


    /**
     * make sure that the T pojo should have used the annotations
     * this method creates a new WorkBook
     *
     * @param listObjs the list of some pojo that we need to write to excel
     * @param <T>      the pojo that has used annotations {@link annotations.ExcelColumn} and {@link annotations.ExcelColumn}
     * @return the created workbook
     */
    @Override
    public <T> Workbook create(List<T> listObjs) throws NoSuchFieldException, IllegalAccessException {
        ExcelCreator<T> excelCreator = new ExcelCreator<T>(listObjs);
        return excelCreator.getWorkBook();
    }


    /**
     * private class that handles the implementation of creating a excel
     */
    private class ExcelCreator<T> {

        private Workbook wb;

        private ExcelCreator(List<T> listObjs) throws NoSuchFieldException, IllegalAccessException {
            createExcel(listObjs);
        }

        Workbook getWorkBook() {
            return wb;
        }

        /**
         * @return object of newly created workbook
         */
        private void createWorkBook() {
            wb = new SXSSFWorkbook();
        }

        /**
         * this methods adds a new sheet to workbook
         *
         * @param sheetName the name of sheet
         */
        private void createSheet(@NotNull String sheetName) {
            Preconditions.checkNotNull(sheetName);
            wb.createSheet(WorkbookUtil.createSafeSheetName(sheetName));
        }


        /**
         * This method adds headers to a particular sheet that is already present in a
         *
         * @param sheetName the name of sheet in which we need to add headers
         * @param headers   the list of headers
         */
        private void writeHeaderstoSheet(@NotNull String sheetName, @NotNull List<String> headers) {
            Preconditions.checkNotNull(sheetName);
            Preconditions.checkNotNull(headers);

            //create a header row and obviously header row is supposed to be at 0
            Row row = wb.getSheet(sheetName).createRow(0);

            Cell cell;
            int cellNum = 0;
            //just styling
            CellStyle cs = getBoldCellStyle();

            //creating header columns
            for (String header : headers) {
                cell = row.createCell(cellNum++);
                cell.setCellStyle(cs);
                cell.setCellValue(header);
            }
        }


        /**
         * @param listObjs the pojo list . The pojo should have implemented {@link annotations.ExcelReport} and {@link annotations.ExcelColumn}
         * @param <T>      the pojo with annotations implemented
         */
        private <T> void createExcel(@NotNull List<T> listObjs)
                throws NoSuchFieldException, IllegalAccessException {

            Preconditions.checkNotNull(listObjs);

            if (listObjs.isEmpty()) {
                throw new IllegalArgumentException("The @param list Objs is empty");
            }

            Class clazz = listObjs.get(0).getClass();
            ExcelReport excelReport = (ExcelReport) clazz.getAnnotation(ExcelReport.class);

            String sheetName = excelReport.value();
            LabelAndFields labelsAndFields = getAnnotedFields(clazz);

            createWorkBook();
            createSheet(sheetName);
            writeHeaderstoSheet(sheetName, labelsAndFields.getLabels());
            writeDatatoSheet(sheetName, listObjs, labelsAndFields.getFields());

        }


        private <T> void writeDatatoSheet(@NotNull String sheetName, @NotNull List<T> listObjs, @NotNull List<Field> fields)
                throws NoSuchFieldException, IllegalAccessException {

            Preconditions.checkNotNull(sheetName, "@param sheetname is null");
            Preconditions.checkNotNull(listObjs, "@param listOjs is null");
            Preconditions.checkNotNull(fields, "@paran fieldNames is null");

            if (listObjs.isEmpty() || fields.isEmpty()) {
                throw new IllegalArgumentException("@param listObjs or @param field name is empty");
            }

            Sheet sheet = wb.getSheet(sheetName);
            int rowNum = 1;//as I expected there will be headers, so 0 will be for headers

            Iterator<T> itr = listObjs.iterator();
            while (itr.hasNext()) {
                int cellNum = 0;
                T t = itr.next();
                Row row = sheet.createRow(rowNum);
                for (Field field : fields) {
                    Cell cell = row.createCell(cellNum++);
                    field.setAccessible(true);// so that we can access private fields
                    cell.setCellValue((String) (field.get(t)));
                }
                rowNum++;
            }
        }


        /**
         * @param clazz the class from which we want to find annotated fields
         */
        private LabelAndFields getAnnotedFields(Class clazz) {
            Field[] fields = clazz.getDeclaredFields();

            List<String> labelList = new ArrayList<String>();
            List<Field> fieldList = new ArrayList<Field>();

            for (Field field : fields) {
                ExcelColumn excelColumn = field.getAnnotation(ExcelColumn.class);
                if (excelColumn != null && !excelColumn.ignore()) {
                    labelList.add(excelColumn.label());
                    fieldList.add(field);
                }
            }

            return new LabelAndFields.Builder()
                    .fields(fieldList)
                    .labels(labelList)
                    .build();
        }

        /**
         * Create CellStyle after setting Font to Bold
         *
         * @return CellStyle
         */
        private CellStyle getBoldCellStyle() {
            Font f = wb.createFont();
            f.setBoldweight(Font.BOLDWEIGHT_BOLD);
            CellStyle cs = wb.createCellStyle();
            cs.setFont(f);
            return cs;
        }
    }
}
