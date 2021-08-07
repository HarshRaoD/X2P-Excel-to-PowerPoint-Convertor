# X2P-Excel-to-PowerPoint-Convertor
A desktop app thats helps create powerpoint presentations using graphs created from Excel data.  

### Input
.csv and .xlsx files can be used for data input.  

### Charts
#### 1) Histogram
Input:- Column with Numerical Data  
Customization:- Graph Title   
#### 2) Violin Chart
Input:- Column with Numerical Data  
Customization:- Graph Title, Colour of Chart  
#### 3) Pie Chart
Input:- Column with Categorical Data  
Customization:- Graph Title, Colour of Chart, Use of Others (Grouping of the smallest categories together such that their combined contribution does not exceed 10% of the total)    
#### 4) Count Chart
Input:- Column with Categorical Data  
Customization:- Graph Title, Colour of Chart  
#### 5) Bar Chart
Input:- 1 Column with Numerical Data and 1 Column with Categorical Data  
Customization:- Graph Title, Colour of Chart, Inversion of X and Y axis, Choice of  sum or mean, Grouped Bar Charts (use of another column with categorical data as Hue)  
#### 6) Line Chart
Input:- 2 Columns with Numerical or Time Series Data  
Customization:- Graph Title, Colour of Chart, Line Design (Solid, Dotted, Dashed, Dotted & Dashed), 2nd Line (use of another column with Numerical or Time Series Data), Line Design for 2nd Line (Solid, Dotted, Dashed, Dotted & Dashed)  
Note:- Column of 2nd Line (if used) must be the same DataType as the First Line .i.e., Numerical and Time Series data cannot be represented at the same time on the same axis  

### Templates
Templates store all chart parameters. They are especially useful for situations where the same data format is used several times. e.g. Weekly/Monthly Sales Review. They can automatically convert the datatype of columns (if requiered) and will ask for permission in case of a forced conversion. There are 3 templates. All of them can be given their own names by the user. Templates dont create the presentation immediately but instead load the data so that changes can be made before creating the presentation. So they can also be used to save time by users exploring the different colours and options available.

### Change of a Column's Data Type
#### 1) To Object/Category (Categorical Data)
Conversion never fails  
#### 2) To Number (Numerical Data)  
If conversion fails, user is asked if he/she wants to convert forcibly (i.e., replace all fields that cant be converted with NaN). Forcible Conversions never fail but may result in loss of data.
#### 3) To Date & Time (Time Series Data)
Conversions can fail and are aborted if they do.

### Title Slide
Users also have the option to add a title and subtitle to the presentation.If provided the Presentation will be saved as <title>.pptx. If no title or subtitle is provided the Presentation will be saved as Presentation_<Date & Time>.pptx with no title slide.  

### Help  
A help button is available in the main window to provide the users with simple instructions to operate the application.
