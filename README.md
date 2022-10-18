# Module 2: VBA of Wall Street

## Overview of Project

### Explain the purpose of this analysis

The purpose of the analysis is to utilize Macro and VBA to organize data. Specifically, we can create a button to receive input from user (year that the user wishes to analyze), and in return, our Macros will output the correct information based on the information chosen (total daily volume and % return based on the year that user has input ).  VBA is further used to group and sum total daily volume based on the "ticker" property using a for loop (add the total volume while the chosen ticker is active). The % return is calculated by finding the starting price (first cell with the selected ticker by checking the cell before) and ending price (last cell with the selected ticker by checking the cell after) using if/then statements. Lastly, the % return is further visualized by assigning different colors for positive (green) and negative (red) return.

## Results - Using images and examples of your code, compare the stock performance between 2018 and 2018, as well as the execution times of the original script and the refactored script.

At a quick glance, for 2017 and 2018 comparison, 2017 has overwhelmingly more positive returns than in 2018 (2017 has 11/12 cells that are green and 2018 has only 2/12 cells that are green). For the original code in 2018, it took much longer to run (0.977 seconds for 2018). On the other hand, refactored script only took 0.125 seconds to run for data in 2018.

## Summary - In a summary statement, address the following questions.
- What are the advantages or disadvantages of refactoring code?

  Advantages of refactoring code include increased organization in coding. The lines of codes are organized efficiently so that it is easier to locate bugs and search for any   potential errors by removing duplicating variables or values. At the end, the program can be run at a much faster pace.

  Disadvantages could include the increased difficulty and complexity to read and write the code. It requires more careful planning and mapping, and more time can be consumed on reworking/refactoring the codes. New errors or bugs can be introduced to the existing code. 

- How do these pros and cons apply to refactoring the original VBA script?

  In this activity, we use arrays (tickersVolume, tickersEndingPrice, and tickersStartingPrice) for the variables to group the values together instead of using a nested for loops to run the same values twice. In return, the program was executed much faster. There's only one variable (tickerIndex) that was declared in the beginning. It is used throughout the code to recall the index of the arrays (tickersVolume, tickersEndingPrice, and tickersStartingPrice) and the values in these arrays are assigned according to the index number as tickerIndex. Since all three arrays only use and depend on one variable (tickerIndex), it eliminates duplication and allows the program to execute much faster. 
In contrast, since three variables are depended on one specific variable (tickerIndex), if there's an error in using tickerIndex, the whole program would not run correctly.


