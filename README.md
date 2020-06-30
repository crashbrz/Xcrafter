![License](https://img.shields.io/badge/license-sushiware-red)
![Issues open](https://img.shields.io/github/issues/crashbrz/Xcrafter)
![GitHub pull requests](https://img.shields.io/github/issues-pr-raw/crashbrz/Xcrafter)
![GitHub closed issues](https://img.shields.io/github/issues-closed-raw/crashbrz/Xcrafter)
![GitHub last commit](https://img.shields.io/github/last-commit/crashbrz/Xcrafter)


# Xcrafter #
Xcrafter is a portable Excel Open XML format command line file crafter. Xcrafter allows you to create xlsx files and embed payloads, like XSS, XXE, SQli,SSRF, and others in an easy and fast way, even without Excel or Calc installed. 

Also, Xcrafter can create regular excel files if you are not looking for a security tool.

### Installation ###
Download the latest release and unpack it in the desired location. Remember to install GoLang in case you want to run from the source.
The Xcrafter uses the Excelize library. Check [https://github.com/360EntSecGroup-Skylar/excelize/](https://github.com/360EntSecGroup-Skylar/excelize/) for more information.

[Here](bin/) you can find the compiled binary.  

### License ###
Xcrafter is licensed under the SushiWare license. Check [docs/license.txt](docs/license.txt) for more information.

### Usage/Help ###
Please refer to the output of -h for usage information and general help. Also, you can contact me on `##spoonfed@freenode.org` (two #)

```
Usage of Xcrafter.exe:
-c string
        Column as a range where the payload will be placed. Use a colon as a range separator and a comma to add a new range. Ex: C1:F1,J7:N7,H1:K1
  -e string
        Use this option to set a different payload from -p option. For single cells only.
  -l string
        Line range of a column where the payload will be placed. Use a colon as a range separator and a comma to add a new range. Ex. A1:A10,B1:B10
  -o string
        Crafted file name output. (default "Xcrafter.xlsx")
  -p string
        Any payload to be written in the file.
  -s string
        Single cells where the payload will be placed. Use a comma as a separator. Ex: A1,H4,D20
  -v    Prints the current version and exit.
  -w string
        Set the worksheet name. (default "Sheet1")
  ```
