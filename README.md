# **Module_2_Challenge**


## **Overview of Project**
### **Purpose**
Previously we created a workbook that would effectively run the same type of analysis over the entire stock market. However, the original coding for the stock analysis tickers had most of the if loops nested in a for loop that resulted in many more iterations over the information to produce the same result. While the difference in time is not significate with this smaller data set, it could be more cumbersome if applied to a larger data set. Our purpose here is to refactor the code to reduce the run time of the stock ticker analysis making it quicker if applied in the future to a larger data set.

### **Results**
The time differences between the refactor of the code can be seen below. The image below show that the refactored code run quite a bit faster than the original code even with the addition of conditional coding.
 
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

1. Refactoring code in general is intended to take code that has already been written and make it function in a way the is more efficient or effective. While the refactored code may end up more efficient or effective the effort put in to rewrite the original may not be worth the end results. Additionally, it is not innately apparent if the new refactored code will truly be truly worth the end results until the code had been interpreted, written, and tested. 

2. The benefit between the original VBA script and the refactored script really depends on the use case scenario. The transparency of the original script makes it much easier to follow and allows a developer to expand upon the string tickers more quickly in less time for an expanded data set with more tickers to track. However, once a developer as had time to understand the way the refactored code is written the refactored code gets the advantage of running at a much quicker rate. The disadvantage both sets of code is time saved in relation to the maintenance needed to modify or maintain the code. Depending on the size of the data set the code would be running through there will be lot more pieces of information to modify in the refactored code than the original, but it may be worth it in the long run if the code is meant to continually and repeatedly run larger data sets - on a smaller set of data the run time benefit is negligible.




