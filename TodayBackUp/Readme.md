# BEE_AHKTests
Automated testing of user interactions with BEE software

# Installation
1-Make sure Excel version is 2010+ \
2- Download and Install ahk from this link https://autohotkey.com/download/ \
3-Download the Repository on your Pc


# Steps to start testing Cygnet
1- UnZib the zip file \
2-unZip the file and open the folder produced from unzipping \
3- (This step and next only for first time) Edit file inside AHKFiles called "Global varible names" change the default location of CygnetTestCasesDefaultFile and CygnetElementsLookupTableDefaultFile by the right destinations
4. Right click on Test application ahk file and click compile script, TestApplication.exe will be created
5- Right click on Test application exe file and choose Run as adminstrator 
## How to Start testing 
1. Change the default excel sheet for test cases or cygnet element excel if you want other wise click Start testing  \
2. Gui will appear, it is to run tests,  you will find at the begining the number of test and seven buttons
    1. the big middle green one is to run all tests 
    2. The two side yellow buttons to either run next or previous test only
    3. In order to Run tests you have first to choose first what test you want to run from the tree by marking the name of the test 
        - Note To know the description of a test you will find it in Test case excel sheet 
            
    4. The two small purple button to choojumb to a previous or next test without running
        - Note 
            1.To run test number 2 choose test number 1 then click either big green button to run tests from 2 till the end or only the yellow next arrow to run only test 2 \
            2. if you run test 0 it will take you to first step which is chossing excel sheets again 
    5. The button on the right of The yellow next icon is checking elements positions to make sure its in the right position or to update the positions
        1. After clicking it, Two Gui-s will open.
            - The first one Called "ScrollGui1" it has all cygnet elements numbered.
            - The second one is to check each element position.
                - "Element index" is the current visible element index
                - Detected image is a screenShot for the current element position
                - Reference image is how the element should look like. Its in the AHKFiles folder - ReferenceApp images
                - At then end of the gui, you will find the description of each button
    
    6. On the right of checking elements positions is console icon to make the console active again in case you closed it, however it is active by default
3. After being done with testing go to the directory you will find new created excelsheet with the same name as TestCases excel sheet plus the date, time of today and Description of each test, open it you will find the latest output of each test you run before
## How to Add new test
1. Inside AHKFiles, create new New test.ahk file, write the test description inside 
2. Add new row in TestCases in the excel sheet with the new test case name, description and position in the tree
3. if you have new Cygnet elements you need for your new test
    - Add there reference images inside "ReferenceApp images" folder
    - Add there names (same image name as in the ReferenceApp images) in Cygnet Elments excel sheet
    - Start testing, then click on change element position
    - jump to your new element 
    - update the position, or click yes to be saved.
4. delete TestApplication shortcut  and TestApplication exe which is inside AHKFiles folder, Then click right click on TestApplication ahk file and choose compile script. New TestApplication exe will be created, you can either start testing using it or create shortcut and put it out of AHK files folder so you will not be bothered with ahk files
5. Predefined functions, you can use
    - CheckCertficate()
    - OpenCygnet()
    -pressMultibleRandomKeys(), Random number of keys from 500-1000 in keyboard 
    - RandomClickNumber(), return random number from 1-10 
    - CloseCygnet()
    - CheckReportTabIsOPen(), check if report tap is open
    - FeedBackIsOPen(), check if feedback tap is open
    - GraphIsOPen(), check if graph tap is open
    - ClickButton(elementName), To click a certain element
    - CheckTab(element), check if element (graph, feedback or report) is open
    - CheckTimerUpdate(), check if Timer is updating
    - closeWindow(windonName), Close certain window
    - writeNewUniqueClient(), write new unque user name
