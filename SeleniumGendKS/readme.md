## Selenium automation testing using C# in NUnit framework.
## Prerequisite:
##### 1. Download and install 3 tools:
 - BuildTools (https://visualstudio.microsoft.com/visual-cpp-build-tools/)
    - (open file 'vs_BuildTools.exe' and install '.NET 6.0 Runtime (LTS)' and '.NET SDK' )
    - (default directory in C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools)
 - Framework64 (https://www.microsoft.com/en-us/download/details.aspx?displaylang=en&id=17851&ppud=4)
    - (default directory in C:\Windows\Microsoft.NET\Framework64)
 - microsoft.testplatform (search at: https://www.nuget.org/packages?q=Microsoft+TestPlatform)
    - (download and rename extension of file to .zip, and then extract to folder -> ex: microsoft.testplatform.17.4.0)

##### 2. How to run bat file:
 - Copy 3 folders above (just installed: `BuildTools`, `Framework64`, `microsoft.testplatform.17.4.0`) to folder path '.../SeleniumGendKS'
 - Check bat file before run:
    - Open file `run_full_testcases.bat` and `run_1_testcase.bat` with Notepad/Notepad++. 
    - At line (13): 
        ```sh
        CD %curdir%\microsoft.testplatform.17.4.0\tools\net451\Common7\IDE\Extensions\TestPlatform
        ```
    - If microsoft testplatform was downloaded with new version (ex: `microsoft.testplatform.17.5.0`) (and `net462`) then correct it:
        ```sh
        CD %curdir%\microsoft.testplatform.17.5.0\tools\net462\Common7\IDE\Extensions\TestPlatform
        ```
 - Check xml config file before run: 
    - The `Config.xml` file is created at folder path `...\Project's name\Config`, to run automation user have to provide an account for this file. --> Please contact the Administrator to have an account.
 - Execute bat files:
    - file `run_full_testcases.bat`: execute bat file to run full testcases
    - file `run_1_testcase.bat`: execute bat file to run 1 testcase. This file is often used to rerun `failed` test cases. (Open this file and copy list test cases from file `run_full_testcases.bat` into this file)