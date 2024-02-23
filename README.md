# Timer Class Module

## Description
The Timer Class Module is a simple tool designed to help detect areas in your own code that might be causing performance issues. It allows you to measure the time taken between different points in your code.

## Author
- **Snowy42**

## Version
- 1.0

## Last Modified
- 10 Feb 2024

## Usage
1. Import the Timer Class Module into your VBA project.
2. Start the timer using the `Start` method.
3. Use the `StoreTime` method at different points in your code to record the time.
4. Call the `Report` method to print the time differences between each recorded point and the start time to the console.

## Example Usage
```vba
Dim myTimer As New TimerClassModule

Sub ExampleUsage()
    myTimer.Start
    ' Your code here
    myTimer.StoreTime "AfterStep1"
    ' More code here
    myTimer.StoreTime "AfterStep2"
    ' Final code here
    myTimer.Report
End Sub
```

## Example Output
```
AfterStep1:   1425.781 ms
AfterStep2:   2832.031 ms
Total Completion Time:      4160.156 ms
```

## Notes
- The reported times are in milliseconds.
- This tool is useful for identifying sections of code that may need optimization.

Feel free to integrate this Timer Class Module into your VBA projects to enhance your code performance analysis.
