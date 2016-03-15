
using System;
using System.Collections.Generic;

using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace Word图表
{
    public partial class WebForm1 : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void Button1_Click(object sender, EventArgs e)
        {
            Aspose.Words.Document doc = new Aspose.Words.Document(); 
            Aspose.Words.DocumentBuilder builder = new Aspose.Words.DocumentBuilder(doc);
            
            // Add chart with default data. You can specify different chart types and sizes.
            Aspose.Words.Drawing.Shape shape = builder.InsertChart(Aspose.Words.Drawing.Charts.ChartType.Column, 432, 252);
            // Chart property of Shape contains all chart related options.
            Aspose.Words.Drawing.Charts.Chart chart = shape.Chart;

           // chart.Title = "csy1111";
            // Get chart series collection.
            Aspose.Words.Drawing.Charts.ChartSeriesCollection seriesColl = chart.Series;
            
            // Delete default generated series.
            
            seriesColl.Clear();
           
            // Create category names array, in this example we have two categories.
            
            string[] categories = new string[] { "第一赛季", "第二赛季","第三赛季" };
            
            // Adding new series. Please note, data arrays must not be empty and arrays must be the same size.
            
            seriesColl.Add("AW Series 1", categories, new double[] { 1, 2 ,2});
            seriesColl.Add("AW Series 2", categories, new double[] { 3, 4 ,44});
            seriesColl.Add("AW Series 3", categories, new double[] { 5, 6,22 });
            seriesColl.Add("AW Series 4", categories, new double[] { 7, 81,3 });
            seriesColl.Add("AW Series 5", categories, new double[] { 9, 10,44 });


            string wordPath = MapPath("/a.docx");
            doc.Save(wordPath);
        }
    }
}