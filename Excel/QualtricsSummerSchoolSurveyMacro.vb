Sub RecreateQualtricsVisualization()
    ' Create Bar Chart Visualization for Q1, Q10 (Matrix Table Question) and Q2, Q3 (Scale of 100 Question)

    ' Q1 - How do you feel about our planned Summer School Curriculum/Programming?
    ' Q2 - On a scale of 0 to 100 (100 being all for it!), how do you feel about every class in our planned Summer School Curriculum?
    ' Q3 - On a scale of 0 to 100 (100 being all for it!), how do you feel about every activity in our planned Summer School Programming?
    ' Q10 - Please choose the time of day that works best for you for dropping off/picking up your child from Summer School:

    Dim BarChartSheetNameArr
    BarChartSheetNameArr = Array("Q1", "Q2", "Q3", "Q10")
    Dim BarChartCellRangeArr
    BarChartCellRangeArr = Array("B3:G5", "B3:H8", "B3:E6", "B3:J5")

    Dim BarChart As ChartObject
    For i = 0 To 3
        Set BarChart = Sheets(BarChartSheetNameArr(i)).ChartObjects.Add(Left:=50, Width:=500, Top:=150, Height:=400)
        With BarChart.Chart
            .SetSourceData Source:=Sheets(BarChartSheetNameArr(i)).Range(BarChartCellRangeArr(i))
            .HasTitle = True
            .ChartTitle.Text = Sheets(BarChartSheetNameArr(i)).Range("A1").Value
            .ChartType = xlBarClustered
        End With
    Next i


    ' Create Pie Chart Visualization for Q4, Q6, Q7 (Fixed Answer Question)

    ' Q4 - What Hands-On Science labs/activities are you interested in having your child participate in?
    ' Q6 - What sports are you interested in having your child participate in?
    ' Q7 - Do you want us to offer the choice of Summer School Lunch for your child?

    Dim PieChartSheetNameArr
    PieChartSheetNameArr = Array("Q4", "Q6", "Q7")
    Dim PieChartCellRangeArr
    PieChartCellRangeArr = Array("B4:C8", "B4:C8", "B4:C7")

    Dim PieChart As ChartObject
    For i = 0 To 2
        Set PieChart = Sheets(PieChartSheetNameArr(i)).ChartObjects.Add(Left:=50, Width:=500, Top:=150, Height:=400)
        With PieChart.Chart
            .SetSourceData Source:=Sheets(PieChartSheetNameArr(i)).Range(PieChartCellRangeArr(i))
            .HasTitle = True
            .ChartTitle.Text = Sheets(PieChartSheetNameArr(i)).Range("A1").Value
            .ChartType = xlPie
            .SetElement msoElementDataLabelInsideEnd
        End With
    Next i


    ' Create Frequency Distribution Visualization for Q5, Q8, Q9 (Text Entry Question)

    ' Q5 - Please make your suggestion for the Summer Reading List for English Language Arts here:
    ' Q8 - What do you want to see added to our Summer School Curriculum?
    ' Q9 - What do you want to see added to our Summer School Programming?

End Sub
