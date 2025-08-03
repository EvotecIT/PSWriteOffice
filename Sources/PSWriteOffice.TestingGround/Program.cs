using PSWriteOffice;
using ShapeCrawler;

namespace PSWriteOffice.TestingGround;

internal class Program {
    static void Main(string[] args) {
        Console.WriteLine("Hello, World!");

        ExampleCreatePresentation();
    }

    public static void ExampleCreatePresentation() {
        // create a new presentation
        var pres = new Presentation();

        var shapes = pres.Slides[0].Shapes;

        // add new shape
        //shapes.AddRectangle(x: 50, y: 60, width: 100, height: 70);
        //var addedShape = shapes.Last();

        //addedShape.TextBox!.Text = "Hello World!";

        //pres.SaveAs("my_pres.pptx");
    }
}
