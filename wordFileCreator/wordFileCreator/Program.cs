using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xceed.Words.NET;

namespace wordFileCreator
{
    class Program
    {
        static void Main(string[] args)
        {
            // region creates the file and inserts the document header
            #region one
            string fileName = "greeting.docx";
            var doc = DocX.Create(fileName);
            doc.InsertParagraph("Hello There!");
            #endregion

            // reates a document header
            #region two
            string title = "What is 12/25 Publishing?";

            //Formats the document header
            // and gives attributes such as font family, font size, 
            // position, font color etc

            Formatting titleFormat = new Formatting();
            titleFormat.FontFamily = new Xceed.Words.NET.Font("Times New Roman");
            titleFormat.Size = 20D;
            titleFormat.Position = 40;
            titleFormat.FontColor = Color.Blue;
            titleFormat.Bold = true;
            titleFormat.UnderlineColor = Color.Black;
            titleFormat.Italic = false;

            // paragraph Text  
            string textParagraph = "That's a good question! " + Environment.NewLine +
                "12/25 Publishing is an independent publishing house with a passion for children's books. " +
                "Please visit Amazon for our latest offering, 'Jeremy Brown and The Unorthodox Crown'" + Environment.NewLine;

            //Formats the paragraph text and sets font family, font size, line spacing, etc  
            Formatting textParagraphFormat = new Formatting(); 
            textParagraphFormat.FontFamily = new Xceed.Words.NET.Font("Times New Roman");
            textParagraphFormat.Size = 12D;
            textParagraphFormat.Spacing = 1;

            //Insert title  
            Paragraph paragraphTitle = doc.InsertParagraph(title, true, titleFormat);
            paragraphTitle.Alignment = Alignment.center;
            //Insert text  
            doc.InsertParagraph(textParagraph, true, textParagraphFormat);
            #endregion

            #region three
            //Create a picture  
            Xceed.Words.NET.Image img = doc.AddImage(@"1225Logo.jpg");
            Picture pic = img.CreatePicture();

            //Create a new paragraph  
            Paragraph par = doc.InsertParagraph("12/25 Publishing Logo. " + Environment.NewLine);

            par.AppendPicture(pic);
            #endregion

            #region four
            //Create Table with 2 rows and 4 columns.  
            Table t = doc.AddTable(4, 2);
            t.Alignment = Alignment.left;
            t.Design = TableDesign.LightList;

            // Fill cells by adding text.  
            // First row
            t.Rows[0].Cells[0].Paragraphs.First().Append("Book Name");
            t.Rows[0].Cells[1].Paragraphs.First().Append("Availability");

            // Second row details
            t.Rows[1].Cells[0].Paragraphs.First().Append("Jeremy Brown and The Unorthodox Crown");
            t.Rows[1].Cells[1].Paragraphs.First().Append("Available");
            t.Rows[2].Cells[0].Paragraphs.First().Append("Jeremy Brown, P.I.");
            t.Rows[2].Cells[1].Paragraphs.First().Append("Available Soon");
            doc.InsertTable(t);
            #endregion

            #region five
            // Hyperlink
            Hyperlink url = doc.AddHyperlink(" Tara King", new Uri("http://itstaraking.wordpress.com"));
            Paragraph p1 = doc.InsertParagraph();
            p1.AppendLine("Please visit the author's site ").Bold().AppendHyperlink(url).Color(Color.Blue); // Hyperlink in blue color 
            #endregion

            #region part of one
            doc.Save();
            Process.Start("WINWORD.EXE", fileName);
            #endregion
        }
    }
}
