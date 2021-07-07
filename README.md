# **Module_2_Challenge**


## **Overview of Project**
### **Purpose**
Previously we created a workbook that would effectively run the same type of analysis over as set of stock market data. However, the original coding for the stock analysis had most of the if loops nested in a for loop that resulted in many more iterations over the information to produce the same result. While the difference in time is not significate with a small data set, it could be more cumbersome if applied to a larger data set. Our purpose here is to refactor the code to reduce the run time of the stock analysis when applied to larger data set.

### **Results**
The run time differences between the refactoring of the code can be seen below. The results show that the refactored code is indeed quicker even with the additional feature of conditional coding added.
 
 ### [Original code](Resources/VBA_challenge_2_original_vba_code.PNG)
 ##### 2017
 ![alt text](https://github.com/kwporras/Module_2_Challenge/blob/5462ce07cf243cfda36b19da3a95064c03f7189e/Resources/VBA_Challenge_2017_Original.PNG)
 
 #### 2018
 ![alt text](https://github.com/kwporras/Module_2_Challenge/blob/5462ce07cf243cfda36b19da3a95064c03f7189e/Resources/VBA_Challenge_2018_Original.PNG)
 
 ### [Refactored code](Resources/VBA_challenge_2_refactored_vba_code.PNG)
 
 #### 2017
 ![alt text](https://github.com/kwporras/Module_2_Challenge/blob/5462ce07cf243cfda36b19da3a95064c03f7189e/Resources/VBA_Challenge_2017.PNG)
 
 ##### 2018
 ![alt text](https://github.com/kwporras/Module_2_Challenge/blob/5462ce07cf243cfda36b19da3a95064c03f7189e/Resources/VBA_Challenge_2018.PNG)
 

### **Summary**

1. Refactoring code in is intended to completed code and make it function either more efficient, effective, or both. While the refactored code can produce better end results the time spend may not justify the time saved. Additionally, it is not innately apparent if refactored code will justify the time spend util the code had been interpreted, written, and tested. 

2. The benefit between the original VBA script and the refactored script depends on the use case scenario. The original VBA script is much easier for different developers to expand upon since the variables are more transparently defined. However, once a developer as had time to understand the way the refactored code is written the original loses some competitive advantage. The disadvantage to both is found in the time saved in relation to time spent on maintenance and modification of code. Depending on the size of the data set the number of factors that need to be updated in the refactored code will cost more time to modify than is saved during the run time. On a smaller set of data, the run time benefit of the refactored code is negligible. However, the refactored code becomes valuable again when larger data sets are running continually and repeatedly.




