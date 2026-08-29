version 18
clear all
set more off

sysuse auto, clear
summarize price mpg weight
regress price mpg weight

display "STATA_GUI_MCP_SMOKE_OK"
