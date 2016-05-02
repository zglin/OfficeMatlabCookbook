# OfficeMatlabCookbook
For those times when you just need to create a chart in Office from Matlab. It goes without saying that this will only work on Win32/Win64 systems.
The COM automation server is mostly a wrapper for the VBA settings in Office.
Many thanks to [Mathworks Support].(http://www.mathworks.com/matlabcentral/answers/99150-is-there-an-example-of-using-matlab-to-create-powerpoint-slides) and to [MSDN](https://msdn.microsoft.com/en-us/library/office/).
## Powerpoint
### Initiating a Powerpobject.
```
   h = actxserver('PowerPoint.Application')
   h.Visible = 1;
```
By making the powerpoint application visible, it's significantly easier to figure out what's going on.

### Poking around - Initial Setup
The invoke command allows you to see what functions you can use with a given object.
```
>> h.Presentations.invoke
	Item = handle Item(handle, Variant)
	Add = handle Add(handle, Variant(Optional))
	Open = handle Open(handle, string, Variant(Optional))
	CheckOut = void CheckOut(handle, string)
	CanCheckOut = bool CanCheckOut(handle, string)
	Open2007 = handle Open2007(handle, string, Variant(Optional))
```
Unless files need to be opened, most of the time, you will be working with the Presentation.Add method.

```
deck=h.Presentations.Add
  
```
If you need to open a file, keep in mind that the open method will open by default in the My-documents folder and if you get a InvokeError stating that PowerPoint could not open the file, you probably forgot to specify the path.
```
filename_path=strcat(pwd,'\',filename)
h.Presentations.Open(filename_path)
```


### Adding Slides
In Office 2003 the com automation server predefined layout templates
```

Slide1 = deck.Slides.Add(1,'ppLayoutBlank')
Slide2 = deck.Slides.Add(2,'ppLayoutBlank')
```
In office 2007 an AddSlide method made in the form of *Addslide(numberofslides,layout)* was introduced. 
layout has to be pulled from the slide master templates
```
>> deck.SlideMaster.CustomLayouts.Item(2).Name
ans =
Title Slide

>>layout = deck.SlideMaster.CustomLayouts.Item(2);
>> Slide1 = deck.Slides.AddSlide(1,layout);

```
If this works, you should see a new slide pop in your powerpoint window.

Most of the time you will want to add a slide after the current slide which necessitates the use of deck.Slides.Count

### Working with Slides
There are two schools of thought around Powerpoint, you can manually build each individual slide and position them, or you can apply formatting to individual templates.
#### Absolute Positioning
Most AddObject Functions take a series of (left edge, top edge, width, height)
```
>>text=Slide1.Shapes.AddTextbox('msoTextOrientationHOrizontal',300,100, 400,400)
>>text.TextFrame.TextRange.Text='Testing text'
```
#### Inserting VBA Charts
While most times you will be ok with a picture of the chart, sometimes there needs to be a dynamic chart

```
>>>chart1=Slide1.Shapes.AddChart(Type, Left, Top, Width, Height)
chart1 =
	Interface.Microsoft_PowerPoint_14.0_Object_Library.Shape
```

This is the default chart, invoke .Chart.ChartData in order to modify existing chart data

##### Changing Chart Titles
Formatting is a bit painful with the VBA object model as there are multiple levels of abstraction. 
To change the title of a slide for example
```
chart1.Chart.ChartTitle.Text= 'sometext'
```
To change the format of the text in the chart itself (for example the axes), is just as involved.
```
chart1.Chart.ChartArea.Format.TextFrame2.TextRange.Font.Size=size
```



### Saving the Presentation 
```
Presentation.SaveAs('C:\...\ExamplePresentation.ppt')
h.quit
h.delete
```
Deleting the server object is important as this will prevent memory issues the next time you use this. 


## Excel 
For Excel, rather than presentations, you would call them workbooks

### Initiating a Workbook
```
   h = actxserver('Excel.Application')
   h.Visible = 1;
```

### Poking Around
...
>> h.Workbooks.invoke
	Add = handle Add(handle, Variant(Optional))
	Close = void Close(handle)
	Item = handle Item(handle, Variant)
	Open = handle Open(handle, string, Variant(Optional))
	OpenText = void OpenText(handle, string, Variant(Optional))
	OpenDatabase = handle OpenDatabase(handle, string, Variant(Optional))
	CheckOut = void CheckOut(handle, string)
	CanCheckOut = bool CanCheckOut(handle, string)
	OpenXML = handle OpenXML(handle, string, Variant(Optional))
...


Create a new workbook
```
>> work=h.Workbooks.Add
 
work =
 
	Interface.Microsoft_Excel_14.0_Object_Library._Workbook
  
```

### Navigating and doing things in workbooks

You should  setup a range variable in order to get Matlab to be able to work with individual cells in matlab. 
...
>> Cursor=activesheet.Range('A1');
 
Cursor =
 
	Interface.Microsoft_Excel_14.0_Object_Library.Range
>> Cursor.Formula=1
 
Cursor =
 
	Interface.Microsoft_Excel_14.0_Object_Library.Range


...
### Running VBA through Workbooks
Macros are run at the excel application level (so whatever variable that is assigned to Excel.Application activex server
...
h.Run('Macroname') 
...
