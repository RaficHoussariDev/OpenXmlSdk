using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using DocumentFormat.OpenXml.Office.Drawing;
using DocumentFormat.OpenXml.Office2010.Word;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Vml;
using DocumentFormat.OpenXml.Wordprocessing;

using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace PowerpointEdit
{
    public class PowerPointEdit
	{
		public PowerPointEdit()
		{
		}

		public void EditPowerPointFile()
		{
			string filePath = "auxi_file.pptx";
			string assemblyLocation = System.Reflection.Assembly.GetExecutingAssembly().Location;

			if (!string.IsNullOrEmpty(assemblyLocation))
			{
                string? directoryPath = System.IO.Path.GetDirectoryName(assemblyLocation);

                if (!string.IsNullOrEmpty(directoryPath))
                {
                    string fullPath = System.IO.Path.Combine(directoryPath, filePath);

                    Console.WriteLine("File Path: " + fullPath);

                    using (PresentationDocument document = PresentationDocument.Open(fullPath, true))
                    {
                        Console.WriteLine("Inside auxi_file.pptx");
                        SlidePart? slidePart = this.GetSlidePartByIndex(document, 0);

                        if (slidePart != null)
                        {
                            // get all the text elements in the slide
                            Console.WriteLine("Slide Part is not null");
                            var textElements = slidePart.Slide.Descendants<A.Text>().ToList();
                            Console.WriteLine("elements count: " + textElements.Count);

                            var firstText = textElements.FirstOrDefault();

                            if(firstText != null)
                            {
                                Console.WriteLine("First Word: " + firstText.Text);
                                firstText.Text = "Output Slide";

                                var paragraphProperties = firstText.Ancestors<A.ParagraphProperties>().FirstOrDefault();
                                if (paragraphProperties == null)
                                {
                                    paragraphProperties = new A.ParagraphProperties();
                                    firstText.InsertBeforeSelf(paragraphProperties);
                                }

                                // Apply font styling (e.g., Typeface) to the paragraph properties
                                var paragraphRunProperties = paragraphProperties.Descendants<A.RunProperties>().FirstOrDefault();
                                if (paragraphRunProperties == null)
                                {
                                    paragraphRunProperties = new A.RunProperties();
                                    paragraphProperties.Append(paragraphRunProperties);
                                }

                                var beirutFont = new A.LatinFont() { Typeface = "Beirut" };
                                paragraphRunProperties.Append(beirutFont);

                                var runProperties = firstText.Descendants<A.RunProperties>().FirstOrDefault();

                                // Modify the run properties if they exist
                                if (runProperties != null)
                                {
                                    // Change color to #FF0000
                                    var solidFill = runProperties.Descendants<SolidFill>().FirstOrDefault();
                                    if (solidFill != null)
                                    {
                                        var srgbColor = solidFill.Descendants<A.RgbColorModelHex>().FirstOrDefault();
                                        if (srgbColor != null)
                                        {
                                            srgbColor.Val = "FFFFF"; // White color
                                        }
                                    }

                                    // Change font size to 48pt
                                    var fontSize = runProperties.Descendants<FontSize>().FirstOrDefault();
                                    if (fontSize != null)
                                    {
                                        fontSize.Val = "4800"; // 48pt in Open XML units
                                    }
                                }
                                else
                                {
                                    Console.WriteLine("null");
                                }

                                //XmlDocument slideXmlDocument = new XmlDocument();
                                //slideXmlDocument.Load(slidePart.GetStream());
                                //XmlNodeList elementsWithClass = slideXmlDocument.SelectNodes("//*[@class]");
                                //Console.WriteLine(elementsWithClass.Count);
                                //foreach (XmlElement element in elementsWithClass)
                                //{
                                //    string classAttribute = element.GetAttribute("class");
                                //    Console.WriteLine(classAttribute);
                                //}

                                //var classAttributes = slideXml.Descendants()
                                //    .Where(element => element.Attribute("class") != null)
                                //    .Select(element => element.Attribute("class").Value)
                                //    .Distinct();

                                // Try to change the text color (Failed)
                                //P.ShapeTree shapeTree = slidePart.Slide.CommonSlideData.ShapeTree;
                                //foreach (var shape in shapeTree.Elements<P.Shape>())
                                //{
                                //    Console.WriteLine(shape.ShapeProperties.InnerText);
                                //    A.TextBody? textBody = shape.Descendants<A.TextBody>().FirstOrDefault();
                                //    if (textBody == null)
                                //    {
                                //        Console.WriteLine("TextBody is null");
                                //    }
                                //}

                                //Append a new Run to the text and change color to white (failed)

                                //var newRun = new A.Run();
                                //var runProperties = new A.RunProperties();
                                //var solidFill = new A.SolidFill();
                                //var rgbColorModelHex = new A.RgbColorModelHex { Val = "FF0000" };
                                //solidFill.Append(rgbColorModelHex);
                                //runProperties.Append(solidFill);
                                //newRun.Append(runProperties);
                                //var newText = new A.Text { Text = "Output Slide" };
                                //newRun.Append(newText);
                                //firstText.Append(newRun);

                                //foreach (A.TextBody textBody in slidePart.Slide.Descendants<A.TextBody>())
                                //{
                                //    string slideXml = textBody.InnerText; // Get the raw XHTML content
                                //    XDocument xDocument = XDocument.Parse(slideXml);

                                //    Console.WriteLine("xdocument: " + xDocument.ToString());

                                //    // Modify the XHTML to set text-align or centering
                                //    foreach (var divElement in xDocument.Descendants("div").Where(div => div.Attribute("class")?.Value == "P6"))
                                //    {
                                //        Console.WriteLine("Here");
                                //        divElement.SetAttributeValue("style", "text-align: center !important;");
                                //    }

                                //    // Set the modified XHTML content back to the slide
                                //    textBody.InnerXml = xDocument.ToString(SaveOptions.DisableFormatting);
                                //}

                                // Apply text alignment (center) to the paragraph properties (Failed)

                                //var alignment = new A.DefaultRunProperties()
                                //{
                                //    TextAlign = Center
                                //};
                                //var alignment = new TextAlignment
                                //{
                                //     = A.TextAlignmentTypeValues.Center,
                                //};

                                //paragraphProperties.Append(alignment);

                                // center the first word
                                //foreach (var divElement in slidePart.Slide.Descendants<Div>())
                                //{
                                //    if (divElement.ClassName != null && divElement.ClassName.Value == "P6")
                                //    {
                                //        // Modify the class attribute to set text alignment to center
                                //        divElement.ClassName.Value = "P10";
                                //    }
                                //}

                                //Console.WriteLine(firstText.OuterXml);

                                //var runProperties = firstText.Descendants<A.RunProperties>().FirstOrDefault();
                                //if (runProperties != null)
                                //{
                                //    var alignment = runProperties.Descendants<Alignment>().FirstOrDefault();
                                //    if (alignment != null)
                                //    {
                                //        alignment.Horizontal = HorizontalAlignmentValues.Center;
                                //    }
                                //    else
                                //    {
                                //        Console.WriteLine("Alignment is null");
                                //    }
                                //}
                                //else
                                //{
                                //    Console.WriteLine("Run Properties is null");
                                //}

                                //var paragraphProperties = firstText.Descendants<A.ParagraphProperties>().FirstOrDefault();
                                //if (paragraphProperties != null)
                                //{
                                //    var alignment = paragraphProperties.Descendants<Alignment>().FirstOrDefault();
                                //    if (alignment != null)
                                //    {
                                //        alignment.Horizontal = HorizontalAlignmentValues.Center;
                                //    }
                                //}
                                //else
                                //{
                                //    Console.WriteLine("Paragraph Properties is null");
                                //}
                            }
                            else
                            {
                                Console.WriteLine("Unable to find first element");
                            }
                        }

                        document.Save();
                    }

                    Console.WriteLine("Document Saved");
                } else
                {
                    Console.WriteLine("Unable to determina the directory path");
                }
            } else
            {
                Console.WriteLine("Unable to determine the assembly location");
            }
		}

		private SlidePart? GetSlidePartByIndex(PresentationDocument document, int index)
		{
			SlideIdList slideIdList = document.PresentationPart.Presentation.SlideIdList;
			SlideId? slideId = slideIdList.ChildElements[index] as SlideId;
			SlidePart? slidePart = document.PresentationPart.GetPartById(slideId?.RelationshipId) as SlidePart;

			return slidePart;	
		}
    }
}

