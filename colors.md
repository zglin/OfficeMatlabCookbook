Colors in office use the .Color.RGB=int property in order to set colors. Unfortunately these colors are built in a crazy way (thanks excel). 

    
    .TextFrame.TextRange.Text='Hello World';
    .TextFrame.TextRange.Font.Size=12;
    .TextFrame.TextRange.ParagraphFormat.Alignment = 'ppAlignCenter';
    .TextFrame.TextRange.Font.Bold = 'msoFalse';
    %Convert RGB colors to an RGB color index
        RGBcolor=[255,0,255]; RGBcolor=rem(RGBcolor,256);
        RGBindex=RGBcolor(1)+RGBcolor(2)*256+RGBcolor(3)*65536;
    .TextFrame.TextRange.Font.Color.RGB=RGBindex
    
