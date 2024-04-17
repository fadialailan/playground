using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;


static void createWordprocessingDocument(string filepath)
{
    // Create a document by supplying the filepath. 
    using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(filepath, WordprocessingDocumentType.Document))
    {

        // Add a main document part. 
        MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();



        // Create the document structure and add some text.
        mainPart.Document = new Document();

        /* mainPart.Document.AppendChild(new PageSize() { Width = 19240, Height = 15840, Orient=PageOrientationValues.Landscape }); */

        Body body = new Body();
        mainPart.Document.AppendChild(body);

        SectionProperties sectionProperties = new SectionProperties();
        body.AppendChild(sectionProperties);
    
        sectionProperties.AppendChild(new PageSize() {Width=12240, Height=15840});

        Paragraph para = body.AppendChild(new Paragraph());

        int paragraph_property_count = para.Elements().Count();

        Run run = para.AppendChild(new Run());
        Run run2 = para.AppendChild(new Run());
        RunProperties run_prop = new RunProperties();
        run_prop.Bold = new Bold();

        new Color() { Val = "FF0000" };



        run.AppendChild(new Text("Create text in body 2 - CreateWordprocessingDocument"));
        run2.Append(run_prop);
        run2.Append(new Text("Bold\nText"));

    }
}

createWordprocessingDocument("test.docx");
