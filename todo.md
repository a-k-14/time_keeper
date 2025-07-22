# Features To Work

## Priority-1
- ~~Popup window to mark tasks active inactive~~
  - ~~small icon button next to drop down to open the popup~~
  - ~~popup is a scrollable frame with two columns in popup > Task, checkbox~~
  - ~~Ok button at the bottom~~ *10-7-25*
- ~~Show daily work hours on the app to avoid opening Excel or Power bi~~ *1-7-25*

## Priority-2
+ Prevent running more than one instance of the app
+ add `debug print` method to get function names and line numbers in the console print
+ ~~rounded corners for the app window~~ *ignored*
+ `refresh button` - to refresh tasklist if any changes are made in Excel manually *(e.g., renamed General to General - Personal while the app was open)*
+ bold column headers in Excel with openpyxl
+ format cell value types for `dates, time` in Timesheet
+ Handle user movement between timezones
+ Account for DST changes
+ Address app crash issues - if the timer is running and crashes, the data of running task is lost

# Issues
+ ~~`system tray` icon isn't cleared after app exit~~ - using quit button newly added to the UI solves this problem 3-7-25
+ Prevet `multi windows` for about and manage tasks