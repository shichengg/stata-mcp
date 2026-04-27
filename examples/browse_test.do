clear all
display "Step 1: start same-do-file session"
sysuse auto, clear
describe

display "Step 2: summary statistics"
summarize price mpg weight
return list

display "Step 3: regression"
regress price mpg weight
ereturn list