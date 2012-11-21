ColumnSelector
==============

Excel 2003+ column selector/multi-view VBA add-in

This is a single VBA Form that can easily be added to almost any add-in project for Excel 2003+. By leveraging the existing Show/Hide column and Comment features of Excel, this Form will allow the user to easily select which columns are to be visible or not, and restore all columns to visibility. It also allows the user to create different views, such that each view is a particular set of columns, and each view can be selected at will, with all other columns being hidden immediately.

The project is fairly intelligent about the existing data set for the spreadsheet, but correct usage simply requires setting a single data area as a named region (or a set of header cells as a named region).

The project can also be used to select rows, and has a few other 'user' (that is, add-in developer) class-based settings, such as whether or not to animate the process of showing/hiding columns.

