using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using System.Drawing.Imaging;
using System.Text.RegularExpressions;
using System.Xml;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

using Newtonsoft.Json.Linq;
using Newtonsoft.Json;

using YLWService;
using YLWService.Extensions;

namespace YLW_WebService.ServerSide
{
    public class RptUtils
    {
        MainDocumentPart _mDoc = null;

        public RptUtils(MainDocumentPart mDoc)
        {
            _mDoc = mDoc;
        }

        public string TableMergeCells(Table oTable, int col1, int col2, int row1, int row2)
        {
            string sRet = "";

            try
            {
                for (int ii = 0; ii < oTable.Elements<TableRow>().Count(); ii++)
                {
                    if (ii < row1 || ii > row2) continue;
                    TableRow oRow = oTable.Elements<TableRow>().ToList()[ii];
                    for (int jj = 0; jj < oRow.Elements<TableCell>().Count(); jj++)
                    {
                        if (jj < col1 || jj > col2) continue;
                        TableCell oCell = oRow.Elements<TableCell>().ToList()[jj];
                        if (ii == row1) oCell.TableCellProperties.VerticalMerge = new VerticalMerge { Val = MergedCellValues.Restart };
                        if (ii > row1 && ii <= row2) oCell.TableCellProperties.VerticalMerge = new VerticalMerge { Val = MergedCellValues.Continue };
                        if (jj == col1) oCell.TableCellProperties.HorizontalMerge = new HorizontalMerge { Val = MergedCellValues.Restart };
                        if (jj > col1 && jj <= col2) oCell.TableCellProperties.HorizontalMerge = new HorizontalMerge { Val = MergedCellValues.Continue };
                    }
                }
            }
            catch (Exception ec)
            {
                sRet = GetMessage(ec.Message, ec.ToString());
            }

            return sRet;
        }

        public string TableMergeCellsH(Table oTable, int rownum, int col1, int col2)
        {
            string sRet = "";

            try
            {
                TableRow oRow = oTable.Elements<TableRow>().ToList()[rownum];
                for (int jj = 0; jj < oRow.Elements<TableCell>().Count(); jj++)
                {
                    TableCell oCell = oRow.Elements<TableCell>().ToList()[jj];
                    if (jj == col1) oCell.TableCellProperties.HorizontalMerge = new HorizontalMerge { Val = MergedCellValues.Restart };
                    if (jj > col1 && jj <= col2) oCell.TableCellProperties.HorizontalMerge = new HorizontalMerge { Val = MergedCellValues.Continue };
                }
            }
            catch (Exception ec)
            {
                sRet = GetMessage(ec.Message, ec.ToString());
            }

            return sRet;
        }

        public string TableMergeCellsV(Table oTable, int colnum, TableRow oRow1, TableRow oRow2)
        {
            string sRet = "";

            try
            {
                int row1 = oTable.Elements<TableRow>().Count();
                int row2 = oTable.Elements<TableRow>().Count();

                for (int ii = 0; ii < oTable.Elements<TableRow>().Count(); ii++)
                {
                    TableRow oRow = oTable.Elements<TableRow>().ToList()[ii];
                    if (oRow == oRow1) row1 = ii;
                    if (oRow == oRow2) row2 = ii;
                }

                return TableMergeCellsV(oTable, colnum, row1, row2);
            }
            catch (Exception ec)
            {
                sRet = GetMessage(ec.Message, ec.ToString());
            }

            return sRet;
        }

        public string TableMergeCellsV(Table oTable, int colnum, int row1, int row2)
        {
            string sRet = "";

            try
            {
                for (int ii = 0; ii < oTable.Elements<TableRow>().Count(); ii++)
                {
                    TableRow oRow = oTable.Elements<TableRow>().ToList()[ii];
                    if (ii < row1 || ii > row2) continue;
                    TableCell oCell = oRow.Elements<TableCell>().ToList()[colnum];
                    if (ii == row1) oCell.TableCellProperties.VerticalMerge = new VerticalMerge { Val = MergedCellValues.Restart };
                    if (ii > row1 && ii <= row2) oCell.TableCellProperties.VerticalMerge = new VerticalMerge { Val = MergedCellValues.Continue };
                }
            }
            catch (Exception ec)
            {
                sRet = GetMessage(ec.Message, ec.ToString());
            }

            return sRet;
        }

        public int RowIndex(Table oTable, TableRow oRow1)
        {
            try
            {
                for (int ii = 0; ii < oTable.Elements<TableRow>().Count(); ii++)
                {
                    TableRow oRow = oTable.Elements<TableRow>().ToList()[ii];
                    if (oRow == oRow1) return ii;
                }
                return -1;
            }
            catch (Exception ec)
            {
                return -1;
            }
        }

        public string TableAddRow(Table oTable, int nBase, int nCount)
        {
            string sRet = "";

            try
            {
                TableRow oRow1 = oTable.Elements<TableRow>().ToList()[nBase];

                for (int i = 0; i < nCount; i++)
                {
                    TableRow oRow2 = new TableRow();
                    oRow2 = (TableRow)oRow1.Clone();
                    oTable.Append(oRow2);
                }
            }
            catch (Exception ec)
            {
                sRet = GetMessage(ec.Message, ec.ToString());
            }

            return sRet;
        }

        public string TableInsertRow(Table oTable, int nBase, int nCount)
        {
            string sRet = "";

            try
            {
                TableRow oRow1 = oTable.Elements<TableRow>().ToList()[nBase];

                for (int i = 0; i < nCount; i++)
                {
                    TableRow oRow2 = new TableRow();
                    oRow2 = (TableRow)oRow1.Clone();
                    oTable.InsertAfter(oRow2, oRow1);
                }
            }
            catch (Exception ec)
            {
                sRet = GetMessage(ec.Message, ec.ToString());
            }

            return sRet;
        }

        public TableRow TableInsertBeforeRow(Table oTable, TableRow oBaseRow)
        {
            try
            {
                TableRow oRow2 = new TableRow();
                oRow2 = (TableRow)oBaseRow.Clone();
                return oTable.InsertBefore(oRow2, oBaseRow);
            }
            catch (Exception ec)
            {
                return null;
            }
        }

        public TableRow TableInsertAfterRow(Table oTable, TableRow oBaseRow)
        {
            try
            {
                TableRow oRow2 = new TableRow();
                oRow2 = (TableRow)oBaseRow.Clone();
                return oTable.InsertAfter(oRow2, oBaseRow);
            }
            catch (Exception ec)
            {
                return null;
            }
        }

        public string TableInsertRows(Table oTable, int nBase, int nAddCnt, int nCount)
        {
            string sRet = "";

            try
            {
                List<TableRow> oRows = oTable.Elements<TableRow>().ToList().GetRange(nBase, nAddCnt);

                for (int i = 0; i < nCount; i++)
                {
                    for (int j = nAddCnt; j > 0; j--)
                    {
                        TableRow oRow2 = new TableRow();
                        oRow2 = (TableRow)oRows[j - 1].Clone();
                        TableRow oRow1 = oTable.Elements<TableRow>().ToList()[nBase + nAddCnt - 1];
                        oTable.InsertAfter(oRow2, oRow1);
                    }
                }
            }
            catch (Exception ec)
            {
                sRet = GetMessage(ec.Message, ec.ToString());
            }

            return sRet;
        }

        public string TableRemoveRow(Table oTable, int nRow)
        {
            string sRet = "";

            try
            {
                TableRow oRow1 = oTable.Elements<TableRow>().ToList()[nRow];
                if (oRow1 != null)
                {
                    oTable.RemoveChild(oRow1);
                }
            }
            catch (Exception ec)
            {
                sRet = GetMessage(ec.Message, ec.ToString());
            }

            return sRet;
        }

        public string TableRemoveRow(Table oTable, TableRow oRow)
        {
            string sRet = "";

            try
            {
                if (oRow != null)
                {
                    oTable.RemoveChild(oRow);
                }
            }
            catch (Exception ec)
            {
                sRet = GetMessage(ec.Message, ec.ToString());
            }

            return sRet;
        }

        public string ReplaceTables(IEnumerable<Table> lstTable, string sKey, string sValue)
        {
            string sRet = "";

            try
            {
                foreach (Table oTable in lstTable)
                {
                    sRet = ReplaceTable(oTable, sKey, sValue);
                    if (sRet != "") return sRet;
                }
                return sRet;
            }
            catch (Exception ec)
            {
                sRet = GetMessage(ec.Message, ec.ToString());
            }

            return sRet;
        }

        public string ReplaceTable(Table oTable, string sKey, string sValue)
        {
            string sRet = "";

            try
            {
                if (oTable.InnerText.IndexOf(sKey, StringComparison.CurrentCultureIgnoreCase) > -1)
                {
                    foreach (TableRow oRow in oTable.Elements<TableRow>())
                    {
                        sRet = ReplaceTableRow(oRow, sKey, sValue);
                        if (sRet != "") return sRet;
                    }
                }
            }
            catch (Exception ec)
            {
                sRet = GetMessage(ec.Message, ec.ToString());
            }

            return sRet;
        }

        public string ReplaceTableRow(TableRow oRow, string sKey, string sValue)
        {
            string sRet = "";

            try
            {
                foreach (TableCell oCell in oRow.Elements<TableCell>())
                {
                    if (oCell.InnerText.IndexOf(sKey, StringComparison.CurrentCultureIgnoreCase) > -1)
                    {
                        foreach (Paragraph para in oCell.Elements<Paragraph>())
                        {
                            if (Utils.IsRtf(sValue))
                                sRet = ReplaceRtf(para, sKey, sValue);
                            else if (Utils.IsHtml2(sValue))
                                sRet = ReplaceHtml(para, sKey, sValue);
                            else
                                sRet = ReplaceText(para, sKey, sValue);
                            if (sRet != "") return sRet;
                        }
                        ReplaceTables(oCell.Elements<Table>(), sKey, sValue);
                    }
                }
            }
            catch (Exception ec)
            {
                sRet = GetMessage(ec.Message, ec.ToString());
            }

            return sRet;
        }
        

        public string TableRowBackcolor(TableRow oRow, string fillcol)
        {
            string sRet = "";

            try
            {
                List<TableCell> cells = oRow.Elements<TableCell>().ToList();
                foreach (TableCell oCell in cells)
                {
                    Shading shadings = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = fillcol };
                    oCell.TableCellProperties.Append(shadings);
                }
            }
            catch (Exception ec)
            {
                sRet = GetMessage(ec.Message, ec.ToString());
            }

            return sRet;
        }

        public string ReplaceHeaderPart(Document doc, string sKey, string sValue)
        {
            string sRet = "";

            try
            {
                foreach (HeaderPart part in _mDoc.HeaderParts)
                {
                    foreach (Paragraph para in part.RootElement.Descendants<Paragraph>())
                    {
                        if (para.InnerText.IndexOf(sKey, StringComparison.CurrentCultureIgnoreCase) > -1)
                        {
                            if (Utils.IsRtf(sValue))
                                sRet = ReplaceRtf(para, sKey, sValue);
                            else if (Utils.IsHtml2(sValue))
                                sRet = ReplaceHtml(para, sKey, sValue);
                            else
                                sRet = ReplaceText(para, sKey, sValue);
                            if (sRet != "") return sRet;
                        }
                    }
                }
                return sRet;
            }
            catch (Exception ec)
            {
                sRet = GetMessage(ec.Message, ec.ToString());
            }

            return sRet;
        }

        public string ReplaceTextAllParagraph(Document doc, string sKey, string sValue)
        {
            string sRet = "";

            try
            {
                foreach (Paragraph para in doc.Body.Elements<Paragraph>())
                {
                    if (para.InnerText.IndexOf(sKey, StringComparison.CurrentCultureIgnoreCase) > -1)
                    {
                        if (Utils.IsRtf(sValue))
                            sRet = ReplaceRtf(para, sKey, sValue);
                        else if (Utils.IsHtml2(sValue))
                            sRet = ReplaceHtml(para, sKey, sValue);
                        else
                            sRet = ReplaceText(para, sKey, sValue);
                        if (sRet != "") return sRet;
                    }
                }
                return sRet;
            }
            catch (Exception ec)
            {
                sRet = GetMessage(ec.Message, ec.ToString());
            }

            return sRet;
        }

        public string ReplaceText(Paragraph para, string sKey, string sValue)
        {
            if (para.InnerText.Trim().Equals("")) return "";

            string sRet = "";

            try
            {
                string sText = para.InnerText;

                if (sText.IndexOf(sKey, StringComparison.CurrentCultureIgnoreCase) > -1)
                {
                    foreach (Run oRun in para.Elements<Run>())
                    {
                        sText = oRun.InnerText;

                        if (sText.IndexOf(sKey, StringComparison.CurrentCultureIgnoreCase) > -1)
                        {
                            sValue = Regex.Replace(sValue, "\r\n", "\n", RegexOptions.IgnoreCase);
                            string[] asValue = sValue.Split(new char[] { '\n' }, StringSplitOptions.None);

                            if (asValue.Length > 1)
                            {
                                Text t = oRun.Elements<Text>().FirstOrDefault();
                                if (t == null)
                                {
                                    t = new Text() { Text = sText, Space = SpaceProcessingModeValues.Preserve };
                                    oRun.Append(t);
                                }
                                t.Text = Regex.Replace(t.Text, sKey, asValue[0], RegexOptions.IgnoreCase);

                                for (int j = 1; j < asValue.Length; j++)
                                {
                                    oRun.Append(new Break());
                                    oRun.Append(new Text() { Text = asValue[j], Space = SpaceProcessingModeValues.Preserve });
                                }
                            }
                            else
                            {
                                Text t = oRun.Elements<Text>().FirstOrDefault();
                                if (t == null)
                                {
                                    t = new Text() { Text = sText, Space = SpaceProcessingModeValues.Preserve };
                                    oRun.Append(t);
                                }
                                t.Text = Regex.Replace(t.Text, sKey, sValue, RegexOptions.IgnoreCase);
                            }
                        }
                    }
                }
            }
            catch (Exception ec)
            {
                sRet = GetMessage(ec.Message, ec.ToString());
            }

            return sRet;
        }

        private static Match IsMatch(IEnumerable<Text> texts, int t, int c, string find)
        {
            int ix = 0;
            for (int i = t; i < texts.Count(); i++)
            {
                for (int j = c; j < texts.ElementAt(i).Text.Length; j++)
                {
                    if (find[ix] != texts.ElementAt(i).Text[j])
                    {
                        return null; // element mismatch
                    }
                    ix++; // match; go to next character
                    if (ix == find.Length)
                        return new Match() { EndElementIndex = i, EndCharIndex = j }; // full match with no issues
                }
                c = 0; // reset char index for next text element
            }
            return null; // ran out of text, not a string match
        }

        /// <summary>
        /// Defines a match result
        /// </summary>
        internal class Match
        {
            /// <summary>
            /// Last matching element index containing part of the search text
            /// </summary>
            public int EndElementIndex { get; set; }
            /// <summary>
            /// Last matching char index of the search text in last matching element
            /// </summary>
            public int EndCharIndex { get; set; }
        }

        public string ReplaceHtml(Paragraph para, string sKey, string sValue)
        {
            if (para.InnerText.Trim().Equals("")) return "";

            string sRet = "";

            try
            {
                if (para.InnerText.IndexOf(sKey, StringComparison.CurrentCultureIgnoreCase) > -1)
                {
                    foreach (Run oRun in para.Elements<Run>())
                    {
                        if (oRun.InnerText.IndexOf(sKey, StringComparison.CurrentCultureIgnoreCase) > -1)
                        {
                            //oRun.RemoveAllChildren();
                            Run run = SetRunHtml(_mDoc, sValue);
                            para.InsertAfter(run, oRun);
                            para.RemoveChild(oRun);
                            //para.ReplaceChild(Run, oRun);
                        }
                    }
                }
            }
            catch (Exception ec)
            {
                sRet = GetMessage(ec.Message, ec.ToString());
            }

            return sRet;
        }

        public string ReplaceRtf(Paragraph para, string sKey, string sValue)
        {
            if (para.InnerText.Trim().Equals("")) return "";

            string sRet = "";

            try
            {
                if (para.InnerText.IndexOf(sKey, StringComparison.CurrentCultureIgnoreCase) > -1)
                {
                    foreach (Run oRun in para.Elements<Run>())
                    {
                        if (oRun.InnerText.IndexOf(sKey, StringComparison.CurrentCultureIgnoreCase) > -1)
                        {
                            //Run run = SetRunRtf(_mDoc, sValue);
                            //para.InsertAfter(run, oRun);
                            //para.RemoveChild(oRun);
                            AltChunk cnk = SetAltChunkRtf(_mDoc, sValue);
                            para.Parent.InsertAfter(cnk, para);
                            para.Remove();
                        }
                    }
                }
            }
            catch (Exception ec)
            {
                sRet = GetMessage(ec.Message, ec.ToString());
            }

            return sRet;
        }

        public string SetText(TableCell oCell, string sKey, string sValue)
        {
            string sRet = "";

            try
            {
                if (oCell.InnerText.IndexOf(sKey, StringComparison.CurrentCultureIgnoreCase) > -1)
                {
                    foreach (Paragraph para in oCell.Elements<Paragraph>())
                    {
                        if (Utils.IsRtf(sValue))
                            sRet = ReplaceRtf(para, sKey, sValue);
                        else if (Utils.IsHtml2(sValue))
                            sRet = ReplaceHtml(para, sKey, sValue);
                        else
                            sRet = ReplaceText(para, sKey, sValue);
                        if (sRet != "") return sRet;
                    }
                }
            }
            catch (Exception ec)
            {
                sRet = GetMessage(ec.Message, ec.ToString());
            }

            return sRet;
        }

        public string SetPlaneText(TableCell oCell, string sValue)
        {
            string sRet = "";

            try
            {
                sValue = Regex.Replace(sValue, "\r\n", "\n", RegexOptions.IgnoreCase);
                string[] asValue = sValue.Split(new char[] { '\n' }, StringSplitOptions.None);

                Paragraph oPara = oCell.Elements<Paragraph>().FirstOrDefault();
                if (oPara == null)
                {
                    oPara = new Paragraph();
                    oCell.Append(oPara);
                }
                oPara.RemoveAllChildren<Run>();
                Run oRun = oPara.Elements<Run>().FirstOrDefault();
                if (oRun == null)
                {
                    oRun = new Run();
                    oPara.Append(oRun);
                }
                if (asValue.Length > 1)
                {
                    Text t = oRun.Elements<Text>().FirstOrDefault();
                    if (t == null)
                    {
                        t = new Text() { Text = asValue[0], Space = SpaceProcessingModeValues.Preserve };
                        oRun.Append(t);
                    }
                    for (int j = 1; j < asValue.Length; j++)
                    {
                        oRun.Append(new Break());
                        oRun.Append(new Text() { Text = asValue[j], Space = SpaceProcessingModeValues.Preserve });
                    }
                }
                else
                {
                    Text t = oRun.Elements<Text>().FirstOrDefault();
                    if (t == null)
                    {
                        t = new Text() { Text = sValue, Space = SpaceProcessingModeValues.Preserve };
                        oRun.Append(t);
                    }
                }
            }
            catch (Exception ec)
            {
                sRet = GetMessage(ec.Message, ec.ToString());
            }

            return sRet;
        }

        public Run SetRunHtml(MainDocumentPart mDoc, string sValue)
        {
            sValue = "<html>" + sValue + "/<html>";
            AlternativeFormatImportPart oChunk;
            oChunk = mDoc.AddAlternativeFormatImportPart(AlternativeFormatImportPartType.Xhtml);
            oChunk.FeedData(sValue.ToStream());

            AltChunk oAltChunk = new AltChunk();
            oAltChunk.Id = mDoc.GetIdOfPart(oChunk);
            return new Run(oAltChunk);
        }

        public AltChunk SetAltChunkHtml(MainDocumentPart mDoc, string sValue)
        {
            sValue = "<html>" + sValue + "/<html>";
            AlternativeFormatImportPart oChunk;
            oChunk = mDoc.AddAlternativeFormatImportPart(AlternativeFormatImportPartType.Xhtml);
            oChunk.FeedData(sValue.ToStream());

            AltChunk oAltChunk = new AltChunk();
            oAltChunk.Id = mDoc.GetIdOfPart(oChunk);
            return oAltChunk;
        }

        public Run SetRunRtf(MainDocumentPart mDoc, string sValue)
        {
            AlternativeFormatImportPart oChunk;
            oChunk = mDoc.AddAlternativeFormatImportPart(AlternativeFormatImportPartType.Rtf);
            oChunk.FeedData(sValue.ToStream());

            AltChunk oAltChunk = new AltChunk();
            oAltChunk.Id = mDoc.GetIdOfPart(oChunk);
            return new Run(oAltChunk);
        }

        public AltChunk SetAltChunkRtf(MainDocumentPart mDoc, string sValue)
        {
            AlternativeFormatImportPart oChunk;
            oChunk = mDoc.AddAlternativeFormatImportPart(AlternativeFormatImportPartType.Rtf);
            oChunk.FeedData(sValue.ToStream());

            AltChunk oAltChunk = new AltChunk();
            oAltChunk.Id = mDoc.GetIdOfPart(oChunk);
            return oAltChunk;
        }

        public void SetImageNull(TableCell oCell, Image sPic, Int64 xpos, Int64 ypos, Int64 widthEmus, Int64 heightEmus)
        {
            if (sPic == null)
            {
                sPic = ((System.Drawing.Image)(Properties.Resources.사진없음));
                SetImage(oCell, sPic, xpos, ypos, (widthEmus < 900000L ? widthEmus : 900000L), (heightEmus < 900000L ? heightEmus : 900000L));
            }
            else
            {
                SetImage(oCell, sPic, xpos, ypos, widthEmus, heightEmus);
            }
        }

        public void SetImage(TableCell oCell, Image sPic, Int64 xpos, Int64 ypos, Int64 widthEmus, Int64 heightEmus)
        {
            ImagePart imagePart = _mDoc.AddImagePart(ImagePartType.Png);
            imagePart.FeedData(sPic.ToStream(ImageFormat.Png));

            oCell.RemoveAllChildren();
            string relationshipId = _mDoc.GetIdOfPart(imagePart);

            var element =
                new Drawing(
                    new DW.Inline(
                    new DW.Extent() { Cx = widthEmus, Cy = heightEmus },
                    new DW.EffectExtent()
                    {
                        LeftEdge = 0L,
                        TopEdge = 0L,
                        RightEdge = 0L,
                        BottomEdge = 0L
                    },
                    new DW.DocProperties()
                    {
                        Id = (UInt32Value)1U,
                        Name = "Picture 1"
                    },
                    new DW.NonVisualGraphicFrameDrawingProperties(
                        new A.GraphicFrameLocks() { NoChangeAspect = true }),
                    new A.Graphic(
                        new A.GraphicData(
                        new PIC.Picture(
                            new PIC.NonVisualPictureProperties(
                            new PIC.NonVisualDrawingProperties()
                            {
                                Id = (UInt32Value)0U,
                                Name = "New Bitmap Image.jpg"
                            },
                            new PIC.NonVisualPictureDrawingProperties()),
                            new PIC.BlipFill(
                            new A.Blip(
                                new A.BlipExtensionList(
                                new A.BlipExtension()
                                {
                                    Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}"
                                })
                                )
                            {
                                Embed = relationshipId,
                                CompressionState =
                                A.BlipCompressionValues.Print
                            },
                            new A.Stretch(
                                new A.FillRectangle())),
                            new PIC.ShapeProperties(
                                new A.Transform2D(
                                new A.Offset() { X = xpos, Y = ypos },
                                new A.Extents() { Cx = widthEmus, Cy = heightEmus }),
                                new A.PresetGeometry(
                                new A.AdjustValueList()
                                )
                                { Preset = A.ShapeTypeValues.Rectangle }))
                        )
                        { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
                    )
                    {
                        DistanceFromTop = (UInt32Value)0U,
                        DistanceFromBottom = (UInt32Value)0U,
                        DistanceFromLeft = (UInt32Value)0U,
                        DistanceFromRight = (UInt32Value)0U
                    });

            Paragraph pg = new Paragraph();

            //Paragraph 가운데 정렬
            ParagraphProperties pgp = new ParagraphProperties();
            pgp.Append(new Justification() {  Val = JustificationValues.Center });
            pg.Append(pgp);

            //TableCell 가운데 정렬
            TableCellProperties tcp = new TableCellProperties();
            tcp.Append(new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center });
            oCell.Append(tcp);

            pg.Append(new Run(element));
            oCell.Append(pg);
        }

        public string GetText(TableCell oCell)
        {
            try
            {
                return oCell.InnerText;
            }
            catch 
            {
                return "";
            }
        }

        public void ReplaceInternalImage(string imageName, Image sPic)
        {
            var imagesToRemove = new List<Drawing>();

            IEnumerable<Drawing> drawings = _mDoc.Document.Descendants<Drawing>().ToList();
            foreach (Drawing drawing in drawings)
            {
                DW.DocProperties dpr = drawing.Descendants<DW.DocProperties>().FirstOrDefault();
                if (dpr != null && dpr.Name == imageName)
                {
                    foreach (A.Blip b in drawing.Descendants<A.Blip>().ToList())
                    {
                        OpenXmlPart imagePart = _mDoc.GetPartById(b.Embed);

                        if (sPic == null)
                        {
                            imagesToRemove.Add(drawing);
                        }
                        else
                        {
                            imagePart.FeedData(sPic.ToStream(ImageFormat.Png));
                        }
                    }
                }

                foreach (var image in imagesToRemove)
                {
                    image.Remove();
                }
            }
        }

        public Table GetTable(IEnumerable<Table> tables, string sKey)
        {
            if (tables == null) return null;
            foreach (Table t in tables)
            {
                if (t.InnerText.IndexOf(sKey, StringComparison.CurrentCultureIgnoreCase) > -1)
                {
                    return t;
                }
            }
            return null;
        }

        public Table GetTable(IEnumerable<Table> tables, params string[] sKey)
        {
            if (tables == null) return null;
            foreach (Table t in tables)
            {
                bool bFindAll = true;
                for (int i = 0; i < sKey.Length; i++)
                {
                    if (t.InnerText.IndexOf(sKey[i], StringComparison.CurrentCultureIgnoreCase) < 0)
                    {
                        bFindAll = false;
                        break;
                    }
                }
                if (bFindAll) return t;
            }
            return null;
        }

        public Table GetSubTable(Table table, string sKey)
        {
            if (table == null) return null;
            foreach (TableRow row in table.Elements<TableRow>())
            {
                if (row.InnerText.IndexOf(sKey, StringComparison.CurrentCultureIgnoreCase) > -1)
                {
                    foreach (TableCell cell in row.Elements<TableCell>())
                    {
                        if (cell.InnerText.IndexOf(sKey, StringComparison.CurrentCultureIgnoreCase) > -1)
                        {
                            return GetTable(cell.Elements<Table>(), sKey);
                        }
                    }
                }
            }
            return null;
        }

        public TableRow GetTableRow(IEnumerable<TableRow> rows, string sKey)
        {
            if (rows == null) return null;
            foreach (TableRow t in rows)
            {
                if (t.InnerText.IndexOf(sKey, StringComparison.CurrentCultureIgnoreCase) > -1)
                {
                    return t;
                }
            }
            return null;
        }

        public TableCell GetTableRow(IEnumerable<TableCell> cells, string sKey)
        {
            if (cells == null) return null;
            foreach (TableCell t in cells)
            {
                if (t.InnerText.IndexOf(sKey, StringComparison.CurrentCultureIgnoreCase) > -1)
                {
                    return t;
                }
            }
            return null;
        }

        public string GetFieldName(string prefix, string fld)
        {
            return "@" + prefix + fld + "@";
        }

        public static void AppendFile(MainDocumentPart mainPart, string filename, bool bNewPage = false)
        {
            if (bNewPage)
            {
                Paragraph para = new Paragraph(new Run((new Break() { Type = BreakValues.Page })));
                mainPart.Document.Body.InsertAfter(para, mainPart.Document.Body.LastChild);
            }

            AlternativeFormatImportPart chunk = mainPart.AddAlternativeFormatImportPart(AlternativeFormatImportPartType.WordprocessingML);
            string altChunkId = mainPart.GetIdOfPart(chunk);

            using (FileStream fileStream = File.Open(filename, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                chunk.FeedData(fileStream);
            }

            AltChunk altChunk = new AltChunk { Id = altChunkId };
            mainPart.Document.Body.AppendChild(altChunk);
        }

        //MergeInNewFile 는 잘 안됨. 
        void MergeInNewFile(string resultFile, IList<string> filenames)
        {
            using (WordprocessingDocument document = WordprocessingDocument.Create(resultFile, WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = document.AddMainDocumentPart();
                mainPart.Document = new DocumentFormat.OpenXml.Wordprocessing.Document(new Body());

                for (int ii = 0; ii < filenames.Count; ii++)
                {
                    string filename = filenames[ii];
                    AlternativeFormatImportPart chunk = mainPart.AddAlternativeFormatImportPart(AlternativeFormatImportPartType.WordprocessingML);
                    string altChunkId = mainPart.GetIdOfPart(chunk);

                    using (FileStream fileStream = File.Open(filename, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    {
                        chunk.FeedData(fileStream);
                    }

                    AltChunk altChunk = new AltChunk { Id = altChunkId };
                    if (ii > 0)
                    {
                        Paragraph para = new Paragraph(new Run((new Break() { Type = BreakValues.Page })));
                        mainPart.Document.Body.InsertAfter(para, mainPart.Document.Body.LastChild);
                    }
                    mainPart.Document.Body.AppendChild(altChunk);
                }

                mainPart.Document.Save();
            }
        }

        public static string GetMessage(string sSbuject, string sDetail)
        {
            return string.Format("[{0}]{1}", sSbuject, sDetail);
        }
    }
}
