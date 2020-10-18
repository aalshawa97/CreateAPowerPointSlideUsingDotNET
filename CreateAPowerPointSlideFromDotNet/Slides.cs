using DocumentFormat.OpenXml.Office2010.ExcelAc;
using DocumentFormat.OpenXml.Presentation;
using System.Collections.Generic;
using System.IO;

namespace CreateAPowerPointSlideFromDotNet
{
    class Slides
    {
        private object presentation;
        private object pptDatei;

        //public Microsoft.Office
        private void AddSlides(string folderName)
        {
            string[] filePaths = Directory.GetFiles(folderName, "*.pptx", SearchOption.TopDirectoryOnly);
            //Microsoft.Office.Core.MsoTrioState oTrue = Microsoft.Office.Core.MsoTriState.msoTrue;
            //Microsoft.Office.Core.MsoTriState oTrue = Microsoft.Office.Core.MsoTriState.msoTrue;
            //Create a presentation
            //Presentation pres = new Presentation();
            //Add the title slide
            //ISlide slide = pres.Slides.AddEmptry(Presentation.LayoutSlide[0]);
            //Set sub title text
            //((IAutoShape.Shapes[0]).TextFrame.Text) = "Slide Title Heading";
            //Open presentation and convert slides
            //presentation.LoadFromFile(@"C:\input.pptx");
            //if (presentation.LoadFromFile(@"C:\input.pptx");
            List<string> texts = new List<string>();
            /*
            for(int i = 0; i < presentation.Slides.Count; i++)
            {

            }
            */
            /*if(pptDatei.Count>0)
            {

            }*/
        }
    }
}
