//Abdullah Mutaz Alshawa
//10/16/2020
//Create a PowerPoint slide from C#.NET
/*
 * Intructions:
Create a solution that a user can use to aid them in generating a PowerPoint slide. The ultimate goal is to create a PowerPoint slide for the user. Keep a user mindset throughout the process and what would best help them. 
They want the solution to give them suggestions of images to use from the internet, based on the contents of the information they are using for the slide. The larger the selection the better, and the preferred minimum is 9 suggestions or more. 
They want to improve efficiency and save time not having to search for images for every slide that they are making for their presentations.

Create a solution that a user can use to aid them in generating a PowerPoint slide. The ultimate goal is to create a PowerPoint slide for the user. Keep a user mindset throughout the process and what would best help them. They want the solution to give them suggestions of images to use from the internet, based on the contents of the information they are using for the slide. The larger the selection the better, and the preferred minimum is 9 suggestions or more. They want to improve efficiency and save time, not having to search for images for every slide that they are making for their presentations.

Create a solution that accepts user input and provides output; 
Title area (input) 
Text area (input) 
image suggestion area (multiple selection); utilize words in the title, and bold words in the text area to bring suggested images in from the internet, with ability to select multiple images (minimum of 3 images available to be put on slide, potentially more) to be included in the slide 

Final solution should use Title, text areas, and using selected images to build a PowerPoint slide. 

Solution should be generated using Windows forms or WPF.

Please note: API's/3rd party API's are not allowed in the solution. Google web hooks can be used instead of 3rd party APIs. (Any time an API is involved, the solution never works for this manager)

Things that can prevent you from “passing”:

-Text in body cannot be bolded and no search function based on text in the body

-Application looks for specific words in title to search for, hard coded, and not dynamic

-3rd party API’s are used

-Pictures come in for title words and no way to bold text in the body

-Missing a way to select up to 3 images
*/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
//using static System.Net.Mime.MediaTypeNames;
//using Word = Microsoft.Office.Interop.Word;
namespace CreateAPowerPointSlideFromDotNet
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            //Welcome the user
            Console.WriteLine("Welcome to PowerPoint from C#!");
            string[] PictureFile = { @"C:\powerpoint\img1.jpg", @"C:\powerpoint\img2.jpg", @"C:\powerpoint\img3.jpg", @"C:\powerpoint\img4.jpg" };
            Application pptApplication;

            //Use the media router to show a presentation
            //MediaRouter mediaRouter = (mediaRouter)ContextBoundObject.getSystemService(ContextBoundObject.MEDIA_ROUTER_SERVICE);

            //Create the presentation file
            //Presentation pptPresentation = pptApplication.Presentations.Add(MsoTriState.msoTrue);
            //Presentation pptpresentation = pptApplication.Presentations.Add();
            for (int i = 0; i < PictureFile.Length; i++)
            {
                //Microsoft.Office.Interop.PowerPoint.Slides slides;

            }
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }
    }
}
