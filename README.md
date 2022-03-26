# stock_analysis

## Overview of Stock Analysis Refactoring Project

### Purpose
Steve required an upgrade from the initial code spanning a single worksheet to one that could be used for multiple worksheets, and he had expressed concerns about the speed at which the code processed the large data sets. The purpose of this project was to meet his requirements for efficacy as well as speed. 

## Analysis and Results

### Analysis of Project
The initial run of the code prior to refactoring had moderate speeds for both the 2017 and 2018 worksheet. The declaration of the worksheet in the opening statement made the code hard to manage as one would have to manually create multiple modules to effectively run an analysis for a different year, or risk re-coding the module any time a different year needed to be chosen.

![Example of Declared Worksheet Prior to Refactored Code](<./resources/specified_year_coding_example.png>)

![2017 Results Prior to Refactored Code](<./resources/prefactor_VBA_Challenge_2017.png>)

![2018 Results Prior to Refactored Code](<./resources/prefactor_VBA_Challenge_2018.png>)

After refactoring, we saw a decrease in the elapsed run time for the code on either worksheet desired. The fluidity Steve required in selection of worksheets was enhanced by assigning an input variable rather than selective coding. 

![Example of Input Variable](<./resources/input_variable_coding.png>)

![2017 Results - Refactored](<./resources/VBA_Challenge_2017.png>)

![2018 Results - Refactored](<./resources/VBA_Challenge_2018.png>)


## Advantages and Disadvantages of Refactoring Code

- The Advantages of Refactoring Code
An advantage in refactoring code is that you do not have to work from scratch and create a solution for a complex issue that may exist outside of your skillset. It is significantly faster to refactor, "stand on the shoulders of giants" as it were, and cite your sources rather than dedicate time you may not have available to researching in order to return a deliverable that meets industry standard and the user specifications. 

- The Disadvantages of Refactoring Code
A significant disadvantage in refactoring code is that every user will write their code differently. If a code snipped is poorly commented, you may see a piece of code as ill-advised and remove it as there is no documentation stating what this code will affect down the line, resulting in errors. You can also be required to refactor code outside of your scope and abilities, or simply within an outdated and ineffective program. 
