# Maximum Moment in the Group
#### Video Demo:  <URL https://youtu.be/5n0cmhXwd4o>
#### Description:
In this project, I tried to make the Python code to find the maximum moment value in each group that was created by the Staad Pro. The Staad Pro software is a structural analysis and design software that can do the analysis of the some structural or some buildings then give us the result as graph or table. In my case, after I have done the analysis phase. I have to spend some time to classifiying the data. So, I wrote the Python program to interact with my structural analysis software (Staad Pro) via the comtypes and ctypes libraries of Python that have the workflow details as follows:

##### Set up and define variables
In this process, I started with setting the operation that can interact between Python and my software via the comtypes library then define some variable which will be used for store the data for the reture value from the API function. Because most of the return value will return the value as variant. So in this part, the 4 types of variable will be defined through the automation function from comtypes for creating double, int, long and string. Next for the first time that I call the API function, It is important to flag the function as method to make the function can be called properly.

##### The AllLoads Function
When we do the analysis of structure, there are so many scenario that we want to do the simulation or analysis. In this function, I captured all of scenrios called load case from 2 different types the combined its together and return a list which contain the number of that scenarios.

##### The AllGroups Function
From the Staad Pro software, normally user will groups some elements that have the same condition into 1 groups for easy to manipulate. But there are many condition. So my code try to check how many of them then try to return the list of the name of these condition which were called group name.

##### The MaxMoFunc Function
There is a value called Moment which is represent how much of the element in the structura carry the force at the situation. So in this function will take 2 arguments from before as inputs then loop to every groups every element and every load cases to find the maximum value of moment then store its in the variable called MomentData before return its.

##### The MakeTable Function
After we got all the information that we need. this function will take the data from previous function then interact to the Staad Pro software to make the software create the table and show it to the user to represent what is element that have maximum moment in each, how much of that maximum moment, what is load case and what is the group name that this element is belong to.
