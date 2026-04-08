namespace TemplateParser.Core;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using System.Reflection;
using DocumentFormat.OpenXml.Bibliography;
using System.Runtime.CompilerServices;
using System.Text.Json;
using DocumentFormat.OpenXml.Office2010.Excel;
using DocumentFormat.OpenXml.Office.CustomUI;

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

            var styleToNodeMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "Heading1", "section" },
                { "Heading2", "subsection" },
                { "Heading3", "subsubsection" },
                // Add more mappings as needed
            };
            var nodeHierarchy = new List<string> { "section", "subsection", "subsubsection", "paragraph" };

            Node root = new Node
            {
                
            };

            Stack<Guid> parentStack = new Stack<Guid>();
            List<Node> nodes = new List<Node>();

            foreach (Paragraph p in body.Descendants<Paragraph>())
            {


                //extracting and displaying the text style
                string?style = p?.ParagraphProperties?.ParagraphStyleId?.Val ?? "No Style";

                //Extracting and displaying the actual text
                string? text = p?.InnerText;

                // 2) [Week 2] Build section hierarchy using Word heading styles.

                //Notes: Maybe use a dictionary to back a new tree data structure,
                // since guid can be used for O(1) parent lookups
                // - currently nodes are stored in a flat list, but there is no special
                // data structure, since parent-child relationships can be determined by
                // parentId.

                //We are supposed to have some sort of recursion:
                // SO we may need to rework the parsing loop to be recursive

                //Despite the short comings of the current approach, it should work
                //all thats left is to assign Json metadata to the node, and
                //HOW DO WE RECONSTRUCT INTO A JSON FILE????

                Guid newNodeId = Guid.NewGuid();
                string nodeType = styleToNodeMap.TryGetValue(style, out string mappedType) ? mappedType : "paragraph";
                string newTitle = text ?? string.Empty;
                
                int hierarchyDiff = nodeHierarchy.IndexOf(nodeType) - (parentStack.Count > 0 ? nodeHierarchy.IndexOf(nodes.Find(n => n.Id == parentStack.Peek()).Type) : -1);
                switch (hierarchyDiff)
                {
                    case 0: // same level
                        if (parentStack.Count > 0) parentStack.Pop();
                        break;
                    case > 0: // lower level, do nothing
                        break;
                    case < 0: // higher level, pop until we find the correct parent
                        while (parentStack.Count > 0 && nodeHierarchy.IndexOf(nodes.Find(n => n.Id == parentStack.Peek()).Type) >= nodeHierarchy.IndexOf(nodeType))
                        {
                            parentStack.Pop();
                        }
                        break;
                }
                // Determine order index among siblings, counts shared parentIds
                int orderIndex = nodes.FindAll(n => n.ParentId == (parentStack.Count > 0 ? parentStack.Peek() : (Guid?)null)).Count();
                // Create and add the new node
                Node node = new Node
                {
                    Id = newNodeId,
                    TemplateId = templateId,
                    Type = nodeType,
                    Title = newTitle,
                    OrderIndex = orderIndex,
                    ParentId = parentStack.Count > 0 ? parentStack.Peek() : (Guid?)null,
                    MetadataJson = "{}" // metadata
                };
                node.MetadataJson = JsonSerializer.Serialize(node);
                nodes.Add(node); // add new node to list, parentId points up the tree
                parentStack.Push(node.Id); //new node becomes the current parent
            }

            
        }
        
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