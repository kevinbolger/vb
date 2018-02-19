# What does this do?

This VBA project was created to automate the updating of charts in a specific usecase.

The end user was storing their data series in a fat, short table. Each data series was represented in a column, with the top row being the identifier.

The user wanted to edit a field on an excel sheet that would update some summary stats in a vlookup and also then update the series on the chart.

The customChart class defines all the attributes and methods needed to carry this out.

The main.bas file is just an example of an instance of the class being created for a chart on the page to update it with the new values each time the user updates the id field value.

The program is slightly overengineered on purpose - just to gain some practice in object oriented programming within vba in Excel. A better solution would be to simply store data more appropriately and use Pivot charts.