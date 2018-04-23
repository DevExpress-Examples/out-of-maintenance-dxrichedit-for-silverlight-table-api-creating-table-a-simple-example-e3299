Imports Microsoft.VisualBasic
Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Net
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Documents
Imports System.Windows.Input
Imports System.Windows.Media
Imports System.Windows.Media.Animation
Imports System.Windows.Shapes
Imports System.IO
Imports DevExpress.XtraRichEdit.API.Native
Imports DevExpress.XtraRichEdit.Utils
Imports System.Reflection


Namespace Walkthrough_Creating_Table
	Partial Public Class MainPage
		Inherits UserControl
		Public Sub New()
			InitializeComponent()
			AddHandler richEditControl1.Loaded, AddressOf richEditControl1_Loaded
		End Sub

		Private Sub richEditControl1_Loaded(ByVal sender As Object, ByVal e As RoutedEventArgs)
			CreateStyles()
		End Sub

		Private Sub btnCreateTable_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
			CreateTable()
			FillTable()
			ApplyHeadingStyle()
		End Sub
		Private Sub CreateTable()

			Dim doc As Document = richEditControl1.Document
			' Clear out the document content
			doc.Delete(richEditControl1.Document.Range)
			' Set up header information
			Dim pos As DocumentPosition = doc.Range.Start
			Dim rng As DocumentRange = doc.InsertSingleLineText(pos, "Silverlight Colors")

			Dim cp_Header As CharacterProperties = doc.BeginUpdateCharacters(rng)
			cp_Header.FontName = "Verdana"
			cp_Header.FontSize = 16
			doc.EndUpdateCharacters(cp_Header)
			doc.InsertParagraph(rng.End)
			doc.InsertParagraph(rng.End)

			' Add the table
			doc.Tables.Add(rng.End, 1, 2, AutoFitBehaviorType.AutoFitToWindow)
			' Format the table
			Dim tbl As Table = doc.Tables(0)

			Try
				tbl.BeginUpdate()

				Dim cp_Tbl As CharacterProperties = doc.BeginUpdateCharacters(tbl.Range)
				cp_Tbl.FontSize = 14
				cp_Tbl.FontName = "Verdana"
				doc.EndUpdateCharacters(cp_Tbl)

				' Insert header caption and format the columns
				doc.InsertSingleLineText(tbl(0, 0).Range.Start, "Name")
				doc.InsertSingleLineText(tbl(0, 1).Range.Start, "Color")
				Dim pp_HeadingSize As ParagraphProperties = doc.BeginUpdateParagraphs(tbl(0, 1).Range)
				pp_HeadingSize.Alignment = ParagraphAlignment.Center
				doc.EndUpdateParagraphs(pp_HeadingSize)

				' Apply a style to the table
				tbl.Style = doc.TableStyles("MyTableGridNumberEight")
			Finally
				tbl.EndUpdate()
			End Try
		End Sub

		Private Sub CreateStyles()
			' Define basic style
			Dim tStyleNormal As TableStyle = richEditControl1.Document.TableStyles.CreateNew()
			tStyleNormal.LineSpacingType = ParagraphLineSpacing.Single
			tStyleNormal.FontName = "Verdana"
			tStyleNormal.Alignment = ParagraphAlignment.Center
			tStyleNormal.Name = "MyTableGridNormal"
			richEditControl1.Document.TableStyles.Add(tStyleNormal)

			' Define Grid Eight style
			Dim tStyleGrid8 As TableStyle = richEditControl1.Document.TableStyles.CreateNew()
			tStyleGrid8.Parent = tStyleNormal
			Dim borders As TableBorders = tStyleGrid8.TableBorders

			borders.Bottom.LineColor = Colors.LightGray
			borders.Bottom.LineStyle = TableBorderLineStyle.Single
			borders.Bottom.LineThickness = 0.75f

			borders.Left.LineColor = Colors.LightGray
			borders.Left.LineStyle = TableBorderLineStyle.Single
			borders.Left.LineThickness = 0.75f

			borders.Right.LineColor = Colors.LightGray
			borders.Right.LineStyle = TableBorderLineStyle.Single
			borders.Right.LineThickness = 0.75f

			borders.Top.LineColor = Colors.LightGray
			borders.Top.LineStyle = TableBorderLineStyle.Single
			borders.Top.LineThickness = 0.75f

			borders.InsideVerticalBorder.LineColor = Colors.LightGray
			borders.InsideVerticalBorder.LineStyle = TableBorderLineStyle.Single
			borders.InsideVerticalBorder.LineThickness = 0.75f

			borders.InsideHorizontalBorder.LineColor = Colors.LightGray
			borders.InsideHorizontalBorder.LineStyle = TableBorderLineStyle.Single
			borders.InsideHorizontalBorder.LineThickness = 0.75f

			tStyleGrid8.CellBackgroundColor = Colors.Transparent
			tStyleGrid8.Name = "MyTableGridNumberEight"
			richEditControl1.Document.TableStyles.Add(tStyleGrid8)

			' Define Headings paragraph style
			Dim pStyleHeadings As ParagraphStyle = richEditControl1.Document.ParagraphStyles.CreateNew()
			pStyleHeadings.Bold = True
			pStyleHeadings.ForeColor = Colors.White
			pStyleHeadings.Name = "My Headings Style"
			richEditControl1.Document.ParagraphStyles.Add(pStyleHeadings)
		End Sub

		Private Sub FillTable()
			' Fill the table with data
			Dim doc As Document = richEditControl1.Document
			Dim tbl As Table = doc.Tables(0)
			Try
				tbl.BeginUpdate()
				tbl.TableCellSpacing = Units.InchesToDocumentsF(0.1f)
				Dim colors = GetStaticPropertyDictionary(GetType(Colors))
				For Each colorPair As KeyValuePair(Of String, Object) In colors
					Dim row As TableRow = tbl.Rows.Append()
					row.HeightType = HeightType.Exact
					row.Height = Units.InchesToDocumentsF(1.0f)
					Dim cell As TableCell = row.FirstCell
					cell.VerticalAlignment = TableCellVerticalAlignment.Center
					doc.InsertSingleLineText(cell.Range.Start, colorPair.Key)
					cell.Next.BackgroundColor = CType(colorPair.Value, Color)
				Next colorPair

			Finally
				tbl.EndUpdate()
			End Try
		End Sub

		Private Sub ApplyHeadingStyle()
			Dim doc As Document = richEditControl1.Document
			Dim tbl As Table = doc.Tables(0)
			For Each cell As TableCell In tbl.Rows.First.Cells
				cell.BackgroundColor = Colors.Black
			Next cell
			Dim pp_Headings As ParagraphProperties = doc.BeginUpdateParagraphs(tbl.Rows.First.Range)
			pp_Headings.Style = doc.ParagraphStyles("My Headings Style")
			doc.EndUpdateParagraphs(pp_Headings)
		End Sub

		Private Sub richEditControl1_DocumentLoaded(ByVal sender As Object, ByVal e As EventArgs)
			CreateStyles()
		End Sub

		Private Sub richEditControl1_EmptyDocumentCreated(ByVal sender As Object, ByVal e As EventArgs)
			CreateStyles()
		End Sub

		Public Shared Function GetStaticPropertyDictionary(ByVal t As Type) As Dictionary(Of String, Object)
			Const flags As BindingFlags = BindingFlags.Static Or BindingFlags.Public Or BindingFlags.NonPublic

			Dim map = New Dictionary(Of String, Object)()
			For Each prop In t.GetProperties(flags)
				map(prop.Name) = prop.GetValue(Nothing, Nothing)
			Next prop
			Return map
		End Function

	End Class

End Namespace
