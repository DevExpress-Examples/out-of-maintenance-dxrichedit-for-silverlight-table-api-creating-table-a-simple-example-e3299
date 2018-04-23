using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Shapes;
using System.IO;
using DevExpress.XtraRichEdit.API.Native;
using DevExpress.Office.Utils;
using System.Reflection;


namespace Walkthrough_Creating_Table
{
    public partial class MainPage : UserControl
    {
        public MainPage()
        {
            InitializeComponent();
            richEditControl1.Loaded += new RoutedEventHandler(richEditControl1_Loaded);
        }

        void richEditControl1_Loaded(object sender, RoutedEventArgs e)
        {
            CreateStyles();
        }

        private void btnCreateTable_Click(object sender, RoutedEventArgs e)
        {
            CreateTable();
            FillTable();
            ApplyHeadingStyle();
        }
        private void CreateTable()
        {

            Document doc = richEditControl1.Document;
            // Clear out the document content
            doc.Delete(richEditControl1.Document.Range);
            // Set up header information
            DocumentPosition pos = doc.Range.Start;
            DocumentRange rng = doc.InsertSingleLineText(pos, "Silverlight Colors");

            CharacterProperties cp_Header = doc.BeginUpdateCharacters(rng);
            cp_Header.FontName = "Verdana";
            cp_Header.FontSize = 16;
            doc.EndUpdateCharacters(cp_Header);
            doc.Paragraphs.Insert(rng.End);
            doc.Paragraphs.Insert(rng.End);

            // Add the table
            doc.Tables.Create(rng.End, 1, 2, AutoFitBehaviorType.AutoFitToWindow);
            // Format the table
            Table tbl = doc.Tables[0];

            try {
                tbl.BeginUpdate();

                CharacterProperties cp_Tbl = doc.BeginUpdateCharacters(tbl.Range);
                cp_Tbl.FontSize = 14;
                cp_Tbl.FontName = "Verdana";
                doc.EndUpdateCharacters(cp_Tbl);

                // Insert header caption and format the columns
                doc.InsertSingleLineText(tbl[0, 0].Range.Start, "Name");
                doc.InsertSingleLineText(tbl[0, 1].Range.Start, "Color");
                ParagraphProperties pp_HeadingSize = doc.BeginUpdateParagraphs(tbl[0, 1].Range);
                pp_HeadingSize.Alignment = ParagraphAlignment.Center;
                doc.EndUpdateParagraphs(pp_HeadingSize);

                // Apply a style to the table
                tbl.Style = doc.TableStyles["MyTableGridNumberEight"];
            }
            finally {
                tbl.EndUpdate();
            }
        }

        private void CreateStyles()
        {
            // Define basic style
            TableStyle tStyleNormal = richEditControl1.Document.TableStyles.CreateNew();
            tStyleNormal.LineSpacingType = ParagraphLineSpacing.Single;
            tStyleNormal.FontName = "Verdana";
            tStyleNormal.Alignment = ParagraphAlignment.Center;
            tStyleNormal.Name = "MyTableGridNormal";
            richEditControl1.Document.TableStyles.Add(tStyleNormal);

            // Define Grid Eight style
            TableStyle tStyleGrid8 = richEditControl1.Document.TableStyles.CreateNew();
            tStyleGrid8.Parent = tStyleNormal;
            TableBorders borders = tStyleGrid8.TableBorders;

            borders.Bottom.LineColor = Colors.LightGray;
            borders.Bottom.LineStyle = TableBorderLineStyle.Single;
            borders.Bottom.LineThickness = 0.75f;

            borders.Left.LineColor = Colors.LightGray;
            borders.Left.LineStyle = TableBorderLineStyle.Single;
            borders.Left.LineThickness = 0.75f;

            borders.Right.LineColor = Colors.LightGray;
            borders.Right.LineStyle = TableBorderLineStyle.Single;
            borders.Right.LineThickness = 0.75f;

            borders.Top.LineColor = Colors.LightGray;
            borders.Top.LineStyle = TableBorderLineStyle.Single;
            borders.Top.LineThickness = 0.75f;

            borders.InsideVerticalBorder.LineColor = Colors.LightGray;
            borders.InsideVerticalBorder.LineStyle = TableBorderLineStyle.Single;
            borders.InsideVerticalBorder.LineThickness = 0.75f;

            borders.InsideHorizontalBorder.LineColor = Colors.LightGray;
            borders.InsideHorizontalBorder.LineStyle = TableBorderLineStyle.Single;
            borders.InsideHorizontalBorder.LineThickness = 0.75f;

            tStyleGrid8.CellBackgroundColor = Colors.Transparent;
            tStyleGrid8.Name = "MyTableGridNumberEight";
            richEditControl1.Document.TableStyles.Add(tStyleGrid8);

            // Define Headings paragraph style
            ParagraphStyle pStyleHeadings = richEditControl1.Document.ParagraphStyles.CreateNew();
            pStyleHeadings.Bold = true;
            pStyleHeadings.ForeColor = Colors.White;
            pStyleHeadings.Name = "My Headings Style";
            richEditControl1.Document.ParagraphStyles.Add(pStyleHeadings);
        }

        private void FillTable()
        {
            // Fill the table with data
            Document doc = richEditControl1.Document;
            Table tbl = doc.Tables[0];
            try {
                tbl.BeginUpdate();
                tbl.TableCellSpacing = Units.InchesToDocumentsF(0.1f);
                var colors = GetStaticPropertyDictionary(typeof(Colors));
                foreach (KeyValuePair<string, object> colorPair in colors) {
                    TableRow row = tbl.Rows.Append();
                    row.HeightType = HeightType.Exact;
                    row.Height = Units.InchesToDocumentsF(1.0f);
                    TableCell cell = row.FirstCell;
                    cell.VerticalAlignment = TableCellVerticalAlignment.Center;
                    doc.InsertSingleLineText(cell.Range.Start, colorPair.Key);
                    cell.Next.BackgroundColor = (Color)colorPair.Value;
                }
            }

            finally {
                tbl.EndUpdate();
            }
        }

        private void ApplyHeadingStyle()
        {
            Document doc = richEditControl1.Document;
            Table tbl = doc.Tables[0];
            foreach (TableCell cell in tbl.Rows.First.Cells) {
                cell.BackgroundColor = Colors.Black;
            }
            ParagraphProperties pp_Headings = doc.BeginUpdateParagraphs(tbl.Rows.First.Range);
            pp_Headings.Style = doc.ParagraphStyles["My Headings Style"];
            doc.EndUpdateParagraphs(pp_Headings);
        }

        private void richEditControl1_DocumentLoaded(object sender, EventArgs e)
        {
            CreateStyles();
        }

        private void richEditControl1_EmptyDocumentCreated(object sender, EventArgs e)
        {
            CreateStyles();
        }

        public static Dictionary<string, object> GetStaticPropertyDictionary(Type t)
        {
            const BindingFlags flags = BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic;

            var map = new Dictionary<string, object>();
            foreach (var prop in t.GetProperties(flags)) {
                map[prop.Name] = prop.GetValue(null, null);
            }
            return map;
        }

    }

}
