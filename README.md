# Reporting-Hierarchy-Macro
Excel file with VBA Macro that creates upward and downward subordinate hierarchy reports from employee headcount data.


## About
1.	This Excel VBA Macro was created to allow for employee reporting structures to be quickly and dynamically reviewed. It essentially takes a data set that lists individuals by a unique identifier (i.e. ID or email address) with their reporting manager and provides the following two kinds of reporting:
    -	The first is upward reporting. You provide an ID or email address and the macro will use a while loop to traverse and record each reporting manager until it reaches an individual with no reporting manger (presumably the CEO). The stop check could easily be changed to a certain ID, email address, etc.
    -	The second is a downward report where the user enters the ID or Email of an individual and indicates the number of reporting tiers they want traversed. The macro then uses a recursive search function to navigate through each level of reporting employees up to the indicated level and print the results to a new worksheet.

## Prerequisites
1.	Microsoft Excel 2016

## Running the Tests
1.	Open the demo file. 
2.	Instructions for the macro hot keys are on the ‘Instructions’ tab.
3.	Hit ‘ctrl + m’ to start the upward reporting prompt
    -	Enter ‘123456’ and click OK
    -	A new reporting sheet will appear called ‘123456_Report’ containing the results
4.	Hit ‘ctrl + d’
    -	Enter ‘123463’ for the employee ID and ‘8’ for the tier listing and click OK
    -	A new reporting sheet will appear called ‘123463_Subordinate_Report’ containing the results

## Author
•	Alex Anduss – All VBA Modules
