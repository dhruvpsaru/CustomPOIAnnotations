package service;

import annotations.ExcelColumn;
import annotations.ExcelReport;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Assert;
import org.junit.Test;
import service.impl.ExcelServiceImpl;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

/**
 * Created by dhruvr on 22/7/15.
 */
public class ExcelServiceImplTest {

    private static final String SHEETNAME = "testSheet";//check for sheetname
    private static final String WORKBOOK_FILE = "/tmp/testDhruv.xls";// for windows do as required also check for permissions

    @Test
    public void createTest() throws IllegalAccessException, NoSuchFieldException, IOException {
        ExcelService service = new ExcelServiceImpl();
        Workbook wb = service.create(createPojoList());

        //you can write it on a stream as you like
        //e.g you can write to file stream to create physical excel file
        //or you can make file downloadable from a servlet
        File file = new File(WORKBOOK_FILE);
        FileOutputStream fileOutputStream = new FileOutputStream(file);
        wb.write(fileOutputStream);//so here we have written the file

        //now we will read that file
        FileInputStream fileRead = new FileInputStream(new File(WORKBOOK_FILE));
        XSSFWorkbook xssfWorkbook = new XSSFWorkbook(fileRead);
        Sheet sheet = xssfWorkbook.getSheetAt(0);

        Assert.assertEquals(sheet.getSheetName(), SHEETNAME);//check sheetname is smame
        Assert.assertEquals(sheet.getRow(0).getPhysicalNumberOfCells(), 4);//as in pojo under we have used four annotations//check number of cells are same
        Assert.assertEquals(sheet.getPhysicalNumberOfRows(), 4);//as list of three pojo and one is header// check number of rows is same

        Row rowHeaders = sheet.getRow(0);
        List<String> headerListAsRetrieved = new ArrayList<String>();
        for (int i = 0; i < rowHeaders.getPhysicalNumberOfCells(); i++) {
            headerListAsRetrieved.add(rowHeaders.getCell(i).getStringCellValue());//as I am storing string
        }
        List<String> expectedHeaders = new ArrayList<String>();
        expectedHeaders.add("User-Name");
        expectedHeaders.add("First-Name");
        expectedHeaders.add("User-Role");
        expectedHeaders.add("Last-Name");
        for (String header : expectedHeaders) {
            Assert.assertTrue(headerListAsRetrieved.contains(header));
        }//so now headers should be equal

    }


    private List<User> createPojoList() {
        List<User> listUser = new ArrayList<User>();
        User user1 = new User("dhruvpsaru", "dhruv", "12213412aswe", "admin", "pal");
        User user2 = new User("freak_yuma", "vector", "yuma_12393e", "user", "scale");
        User user3 = new User("testing_33", "testfirstName", "12213412aado", "admin", "testlastname");
        listUser.add(user1);
        listUser.add(user2);
        listUser.add(user3);
        return listUser;
    }


    @ExcelReport(SHEETNAME)
    public class User {

        @ExcelColumn(label = "User-Name")
        private String username;

        @ExcelColumn(label = "First-Name")
        private String firstName;

        //ah I dont want to include it
        private String passHint;

        @ExcelColumn(label = "User-Role")
        private String userRole;

        @ExcelColumn(label = "Last-Name")
        private String lastName;


        public User(String userName, String firstName, String passHint, String userRole, String lastName) {
            this.username = userName;
            this.firstName = firstName;
            this.passHint = passHint;
            this.userRole = userRole;
            this.lastName = lastName;
        }


        //there is no need of getters setters just including them
        //if you comment still everything will work fine

        public void setUsername(String username) {
            this.username = username;
        }

        public void setFirstName(String firstName) {
            this.firstName = firstName;
        }

        public void setPassHint(String passHint) {
            this.passHint = passHint;
        }

        public void setUserRole(String userRole) {
            this.userRole = userRole;
        }

        public void setLastName(String lastName) {
            this.lastName = lastName;
        }

        public String getUsername() {
            return username;
        }

        public String getFirstName() {
            return firstName;
        }

        public String getPassHint() {
            return passHint;
        }

        public String getUserRole() {
            return userRole;
        }

        public String getLastName() {
            return lastName;
        }
    }


}
