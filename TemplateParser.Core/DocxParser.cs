namespace TemplateParser.Core;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using System.Reflection;
using DocumentFormat.OpenXml.Office2010.PowerPoint;
using DocumentFormat.OpenXml.Bibliography;
using System.Text;
using System.Text.Json;
using DocumentFormat.OpenXml.Presentation;

public sealed class DocxParser
{
    public ParserResult ParseDocxTemplate(string filePath, Guid templateId)
    {
        // TODO (Week 1-4): Implement core DOCX parsing here.
        // Recommended responsibilities for this method:
        // 1) [Week 1] Learn DOCX structure and print paragraphs from the document.
        using (WordprocessingDocument wordprocessingDocument = WordprocessingDocument.Open(filePath, false))
        {
            Body? body = wordprocessingDocument?.MainDocumentPart?.Document?.Body;
            ArgumentNullException.ThrowIfNull(body, "Document is empty.");

            //[Week 2] Build section hierarchy using Word heading styles.
            //holding child and parent with stack

            //mapping the style to node type
            var styleToNodeType = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                {"Heading1", "section"},
                {"Heading2", "subsection"},
                {"Heading3", "subsubsection"}
            };

            Stack<Guid> parentStack = new Stack<Guid>();
            List<Node> nodes = new List<Node>();

            foreach (Paragraph p in body.Descendants<Paragraph>())
            {
                // Extracting actual text
                string? text = p?.InnerText;
                //Extracting the style
                string?style = p?.ParagraphProperties?.ParagraphStyleId?.Val ?? "No Style";

                //Creating a new node for each paragraph
                Node node = new Node
                {
                    Id = Guid.NewGuid(),
                    TemplateId = templateId,
                    Type = styleToNodeType.TryGetValue(style, out string nodeType) ? nodeType : "paragraph",
                    Title = text,
                    //if type is same as current parrent, then pop
                    //if type is lower in hier then push
                    //if type is higher in hier then pop until type is same, then push
                    OrderIndex = nodes.FindAll(n => n.Type == node.Type).Count + 1,
                    ParentId = parentStack.Count > 0 ? parentStack.Peek() : (Guid?)null,
                    MetadataJson = "{}"
                };
                //enque node to tree
                nodes.Add(node);
                parentStack.Push(node.Id);
            }
        
        // 2) [Week 2] Build section hierarchy using Word heading styles.
        //TO DO LIST: 
            //--Mapping style names to node types
            //---[heading 1 --> section 
            //    heading 2 --> subsection
            //     heading 3 --> subsubsection] (JSON format)



            //Generating random UUIDs [Random(Guid.NewGuid())]
            // if (string.IsNullOrEmpty(filePath))
            // {
            //     throw new ArgumentException("Document is empty");
            // }
            // //converting ID into bytes
            // byte[] idBytes = templateId.ToByteArray();
            // //converting filepath into bytes
            // byte[] fileBytes = Encoding.UTF8.GetBytes(filePath);


        //

        // 3) [Week 3] Detect tables, lists, and images as structured content nodes.
        // 4) [Week 4] Add formatting heuristics for files missing heading styles.
        // 5) [Week 2-4] Create Node instances with:
        //    - Id: new Guid for each node
        //    - TemplateId: the templateId argument
        //    - ParentId: null for root nodes, set for child nodes
        //    - Type/Title/OrderIndex/MetadataJson based on parsed content
        // 6) [Week 4] Return ParserResult with Nodes in deterministic order.
        //
        // Helper guidance [Week 3-6]:
        // - YES, create helper classes if this method gets long or hard to read.
        // - Keep helpers inside TemplateParser.Core (for example, Parsing/ or Utilities/ folders).
        // - Keep this method as the high-level orchestration entry point.
        // - In Week 6, refactor large blocks from this method into focused helper classes.
        //
        // Do not place parsing logic in the CLI project; keep it in Core.
        throw new NotImplementedException("DOCX parsing is intentionally not implemented in this starter repository.");
        }
    }
}
