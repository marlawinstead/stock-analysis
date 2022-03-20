# stock-analysis
Using Visual Basic for Applications to analyze stock market trends in green energy companies

VBA Challenge

Overview of the Project

In this project, we’re analyzing green energy stocks to determine the best company to invest based on their performance during the years 2017 and 2018. Data will be analyzed through Visual Basic for Applications, or VBA. Using VBA, code is written in the form of macros or subroutines to automate certain Excel formulas. These formulas will help yield data to show which stocks are best to invest in based on their successes.
The purpose of this project is to assist Steve, a finance graduate, as he invests his parents’ money in the best green energy stock. The stock they have originally chosen, DAQO, did not perform well in 2017 and 2018. Steve would like to analyze all green energy companies and give his parents financial advice on which company is the best to invest in based on our analysis.

Results

After refactoring the subroutine, the code runs faster. It performs the same function, but only loops through the data once instead of looping through the data based on the number of stocks present in the spreadsheet. Below are the runtimes of the code and the refactored code.

Original Code:

![Original 2017](https://user-images.githubusercontent.com/100978922/159177497-4409f181-ad83-4786-8b96-49ab0e4819de.png)
![Original 2018](https://user-images.githubusercontent.com/100978922/159177503-47e79055-dcb0-4745-b423-f06d858059ff.png)

Refactored Code:

![Refactored 2017](https://user-images.githubusercontent.com/100978922/159177515-5c304d44-e463-4ba2-bf1f-338b3563a619.png)
![Refactored 2018](https://user-images.githubusercontent.com/100978922/159177517-85d70616-3670-458a-83a1-f7b1e9d2a04d.png)

Summary

The advantages to refactoring the code is that it runs at a significantly faster pace, therefore making it more efficient. While the runtimes shown above are both relatively short, in a much longer spreadsheet, we would see a wider gap in the length of time to obtain the results.

I am unable to find a disadvantage to the refactored code. Both the refactored and original code are highly dependent upon the order of the stocks in the worksheet. If the ordering changed, both the refactored and original code would not run properly. An additional refactoring would be to update the code to make the macro more resilient. 
