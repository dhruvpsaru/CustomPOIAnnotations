I wanted to have a POC on custom annotations and I was working on apache poi for excel export. I wanted to use custom annotations , so on googling I came accross [http://persistentdesigns.com/wp/2010/02/excel-report-via-simple-java-annotations/](http://persistentdesigns.com/wp/2010/02/excel-report-via-simple-java-annotations/). Thanks to him<br />

<br />
 Following this I created a POC on custom annotations, reflections apache POI. In the end I was able to create a generic method that accepts List<any pojo>.<br /> Just pojo should have implemeted my
 cusmtom annotations.

In test Case I have created a excel and created a sample pojo there
 So you can have  at class level 
        
        @ExcelReport(SHEETNAME)
         
            This will create a excel with name as in it
                

                
Other annotaions are
         
          @ExcelColumn(label = "User-Role")
                 private String userRole;
                 
                 This is field level annotaion. It tells what will be header name

                 
          @ExcelColumn(ignore = true)
                 private String userRole;
                 
                 I have not particularly used it on testing. But as implementation if you don't want to have a field in excel.
                 So can either annotate it like this or don't annotate it
                 
                 
                 

                 
  
So right now I am still working on it improving it. I want to have improve it on the so that it works well for apache poi.
Right now I am thinking of setting some priority base thing as annotation.
<br />
Any suggestions will be appreciated. Please reach me on dhruvpsaru@gmail.com