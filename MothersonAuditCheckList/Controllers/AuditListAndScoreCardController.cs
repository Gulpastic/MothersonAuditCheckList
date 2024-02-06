using System.IO;
using Microsoft.AspNetCore.Http.HttpResults;
using Microsoft.AspNetCore.Mvc;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using MothersonAuditCheckList.Models.DTO;
using Org.BouncyCastle.Utilities;
using NPOI.OpenXmlFormats.Spreadsheet;
using System.Collections.Generic;
using NPOI.Util;
using System.Data;


namespace MothersonAuditCheckList.Controllers
{
    public class AuditListAndScoreCardController : ControllerBase
    {
        [HttpPost]
        [ProducesResponseType(StatusCodes.Status200OK)]
        [ProducesResponseType(StatusCodes.Status400BadRequest)]
        [ProducesResponseType(StatusCodes.Status500InternalServerError)]
        [Route("api/CreateExcel")]
        public IActionResult createExcel([FromBody] AuditListDTO auditListDTO)
        {
            #region Create excel workbook and Sheets

            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet checklist = (XSSFSheet)wb.CreateSheet("Safety Audit Checklist");
            XSSFSheet card = (XSSFSheet)wb.CreateSheet("Score Card");

            #endregion

            #region Create logo Header Cell Style and Font

            var Logo = wb.CreateCellStyle();
            Logo.Alignment = HorizontalAlignment.Center;
            Logo.WrapText = true;
            IFont logoFont = wb.CreateFont();
            logoFont.FontName = "Calibri";
            logoFont.FontHeightInPoints = 11;
            Logo.SetFont(logoFont);
            Logo.BorderLeft = BorderStyle.Medium;
            Logo.BorderBottom = BorderStyle.Medium;
            Logo.BorderTop = BorderStyle.Medium;

            #endregion

            #region Create 1st Header Cell Style and Font

            var Header1 = wb.CreateCellStyle();
            Header1.Alignment = HorizontalAlignment.Center;
            IFont HeaderFont1 = wb.CreateFont();
            HeaderFont1.Boldweight = (short)FontBoldWeight.Bold;
            HeaderFont1.FontName = "Calibri";
            HeaderFont1.FontHeightInPoints = 12;
            Header1.SetFont(HeaderFont1);
            Header1.WrapText = true;
            Header1.BorderBottom = BorderStyle.Medium;
            Header1.BorderTop = BorderStyle.Medium;

            #endregion

            #region Create 2nd Header Cell Style and Font

            var Header2 = wb.CreateCellStyle();
            Header2.Alignment = HorizontalAlignment.Center;
            IFont HeaderFont2 = wb.CreateFont();
            HeaderFont2.Boldweight = (short)FontBoldWeight.Bold;
            HeaderFont2.FontName = "Calibri";
            HeaderFont2.FontHeightInPoints = 12;
            Header2.SetFont(HeaderFont2);
            Header2.WrapText = true;
            Header2.BorderRight = BorderStyle.Medium;
            Header2.BorderBottom = BorderStyle.Medium;
            Header2.BorderTop = BorderStyle.Medium;

            #endregion

            #region Create Header Cell Style and Font

            var Header = wb.CreateCellStyle();
            Header.Alignment = HorizontalAlignment.Center;
            IFont HeaderFont = wb.CreateFont();
            HeaderFont.Boldweight = (short)FontBoldWeight.Bold;
            HeaderFont.FontName = "Calibri";
            HeaderFont.FontHeightInPoints = 12;
            Header.SetFont(HeaderFont);
            Header.WrapText = true;
            Header.BorderLeft = BorderStyle.Medium;
            Header.BorderRight = BorderStyle.Medium;
            Header.BorderBottom = BorderStyle.Medium;
            Header.BorderTop = BorderStyle.Medium;

            #endregion

            #region Create checkpointHeader Style

            var checkpointHeader = wb.CreateCellStyle();
            checkpointHeader.Alignment = HorizontalAlignment.Center;
            checkpointHeader.FillForegroundColor = HSSFColor.LightYellow.Index;
            checkpointHeader.FillPattern = FillPattern.SolidForeground;
            IFont checkpointHeaderFont = wb.CreateFont();
            checkpointHeaderFont.Boldweight = (short)FontBoldWeight.Bold;
            checkpointHeaderFont.FontName = "Calibri";
            checkpointHeaderFont.FontHeightInPoints = 12;
            checkpointHeader.SetFont(checkpointHeaderFont);
            checkpointHeader.BorderLeft = BorderStyle.Medium;
            checkpointHeader.BorderRight = BorderStyle.Medium;
            checkpointHeader.BorderBottom = BorderStyle.Medium;
            checkpointHeader.BorderTop = BorderStyle.Medium;
            checkpointHeader.WrapText = true;

            #endregion

            #region Data Cell Style

            var Data = wb.CreateCellStyle();
            Data.Alignment = HorizontalAlignment.Center;
            IFont dataFont = wb.CreateFont();
            dataFont.FontName = "Calibri";
            dataFont.FontHeightInPoints = 11;
            Data.SetFont(dataFont);
            Data.BorderLeft = BorderStyle.Medium;
            Data.BorderRight = BorderStyle.Medium;
            Data.BorderBottom = BorderStyle.Medium;
            Data.BorderTop = BorderStyle.Medium;
            Data.WrapText = true;

            #endregion

            #region add logo on top
            checklist.AddMergedRegion(new CellRangeAddress(1, 1, 1, 2));

            byte[] data = System.IO.File.ReadAllBytes("logo.png");
            int pictureIndex = wb.AddPicture(data, PictureType.PNG);
            XSSFDrawing drawing = (XSSFDrawing)checklist.CreateDrawingPatriarch();
            XSSFClientAnchor anchor = new XSSFClientAnchor();
            anchor.Row1 = 1;
            anchor.Col1 = 1;
            IPicture picture = drawing.CreatePicture(anchor, pictureIndex);
            picture.Resize();
            int pictWidthPx = picture.GetImageDimension().Width;

            float cellWidthPx = 0f;
            for (int col = 1; col < 2; col++)
            {
                cellWidthPx += checklist.GetColumnWidthInPixels(col);
            }

            int centerPosPx = (int)Math.Round(cellWidthPx / 2f);

            int anchorCol1 = 1;
            for (int col = 1; col < 2; col++)
            {
                if (Math.Round(checklist.GetColumnWidthInPixels(col)) < centerPosPx)
                {
                    centerPosPx -= (int)Math.Round(checklist.GetColumnWidthInPixels(col));
                    anchorCol1 = col + 1;
                }
                else
                {
                    break;
                }
            }

            anchor.Col1 = anchorCol1;
            anchor.Dx1 = centerPosPx * Units.EMU_PER_PIXEL;
            var row = checklist.CreateRow(1);

            double scaleX = 1.0;
            double scaleY = 1.0;
            if (picture.GetImageDimension().Width > cellWidthPx)
            {
                scaleX = cellWidthPx / picture.GetImageDimension().Width;
            }

            if (picture.GetImageDimension().Width > row.HeightInPoints)
            {
                scaleY = row.HeightInPoints / picture.GetImageDimension().Height * 1.34;
            }

            anchor.AnchorType = AnchorType.MoveDontResize;
            picture.Resize(scaleX, scaleY);

            #endregion

            #region Create and Merge Cells

            IRow headerRow = checklist.CreateRow(1);
            var headerCell = headerRow.CreateCell(1);
            headerCell.CellStyle = Logo;

            var headerCell1st = headerRow.CreateCell(2);
            headerCell1st.CellStyle = Header1;

            var headerCell2nd = headerRow.CreateCell(3);
            headerCell2nd.CellStyle = Header2;

            XSSFRichTextString richString = new XSSFRichTextString("COSA Safety Audit Checksheet-" + DateTime.Now.ToString("yyyy") + Environment.NewLine + "COMPREHENSIVE CHECKLIST OF ENVIRONMENT, HEALTH & SAFETY AUDIT");
            headerCell2nd.SetCellValue(richString);

            checklist.AddMergedRegion(new CellRangeAddress(1, 1, 3, 6));
            checklist.AddMergedRegion(new CellRangeAddress(2, 2, 2, 3));
            checklist.AddMergedRegion(new CellRangeAddress(3, 3, 2, 3));


            for (int i = 4; i < 7; i++)
            {
                headerCell = headerRow.CreateCell(i);
                headerCell.CellStyle = Header;
            }

            #endregion

            #region 2nd Row Table Contents

            var cellIndex = 1;

            IRow headerRow2 = checklist.CreateRow(2);
            ICell headerCell2 = headerRow2.CreateCell(cellIndex);
            headerCell2.SetCellValue("Unit");
            headerCell2.CellStyle = Header;

            cellIndex = cellIndex + 1;
            headerCell2 = headerRow2.CreateCell(cellIndex);
            headerCell2.SetCellValue(auditListDTO.Unit);
            headerCell2.CellStyle = Header;

            cellIndex = cellIndex + 1;
            headerCell2 = headerRow2.CreateCell(cellIndex);
            headerCell2.CellStyle = Header;

            cellIndex = cellIndex + 1;
            headerCell2 = headerRow2.CreateCell(cellIndex);
            headerCell2.CellStyle = Header;

            cellIndex = cellIndex + 1;
            headerCell2 = headerRow2.CreateCell(cellIndex);
            headerCell2.SetCellValue("Auditors");
            headerCell2.CellStyle = Header;

            cellIndex = cellIndex + 1;
            headerCell2 = headerRow2.CreateCell(cellIndex);
            headerCell2.SetCellValue(auditListDTO.auditors);
            headerCell2.CellStyle = Header;

            #endregion

            #region 3rd Row Contents

            cellIndex = 1;

            IRow headerRow3 = checklist.CreateRow(3);
            ICell headerCell3 = headerRow3.CreateCell(cellIndex);
            headerCell3.SetCellValue("Audit Date");
            headerCell3.CellStyle = Header;

            cellIndex = cellIndex + 1;
            headerCell3 = headerRow3.CreateCell(cellIndex);
            headerCell3.SetCellValue(DateTime.Now.ToString("dd-MM-yyyy"));
            headerCell3.CellStyle = Header;

            cellIndex = cellIndex + 1;
            headerCell3 = headerRow3.CreateCell(cellIndex);
            headerCell3.CellStyle = Header;

            cellIndex = cellIndex + 1;
            headerCell3 = headerRow3.CreateCell(cellIndex);
            headerCell3.CellStyle = Header;

            cellIndex = cellIndex + 1;
            headerCell3 = headerRow3.CreateCell(cellIndex);
            headerCell3.SetCellValue("Auditees");
            headerCell3.CellStyle = Header;

            cellIndex = cellIndex + 1;
            headerCell3 = headerRow3.CreateCell(cellIndex);
            headerCell3.SetCellValue(auditListDTO.auditees);
            headerCell3.CellStyle = Header;

            #endregion

            #region 4th Row Contents

            cellIndex = 1;

            IRow headerRow4 = checklist.CreateRow(4);
            ICell headerCell4 = headerRow4.CreateCell(cellIndex);
            headerCell4.SetCellValue("Section / Rule / Form");
            headerCell4.CellStyle = Header;

            cellIndex = cellIndex + 1;
            headerCell4 = headerRow4.CreateCell(cellIndex);
            headerCell4.SetCellValue("Type");
            headerCell4.CellStyle = Header;

            cellIndex = cellIndex + 1;
            headerCell4 = headerRow4.CreateCell(cellIndex);
            headerCell4.SetCellValue("Sr. No.");
            headerCell4.CellStyle = Header;

            cellIndex = cellIndex + 1;
            headerCell4 = headerRow4.CreateCell(cellIndex);
            headerCell4.SetCellValue("Audit Checkpoints");
            headerCell4.CellStyle = Header;

            cellIndex = cellIndex + 1;
            headerCell4 = headerRow4.CreateCell(cellIndex);
            headerCell4.SetCellValue("Score");
            headerCell4.CellStyle = Header;

            cellIndex = cellIndex + 1;
            headerCell4 = headerRow4.CreateCell(cellIndex);
            headerCell4.SetCellValue("Auditor Remark");
            headerCell4.CellStyle = Header;

            #endregion

            for (int i = 1; i < 4; i++)
            {
                checklist.AutoSizeRow(i);
                for (int j = 1; j < 7; j++)
                    checklist.AutoSizeColumn(j);
            }

            #region Calling body content input & format function

            WriteBodyContent(auditListDTO.RuleList, checklist, checkpointHeader, Data);

            #endregion

            #region Creating Excel File

            string FileName = "Motherson_Audit_Checklist_" + DateTime.Now.ToString("yyyy-dd-MM--HH-mm-ss") + ".xlsx";
            using (FileStream file = new FileStream(@"C:\Users\kartik\OneDrive - Sparrow Risk Management Pvt. Ltd\Documents\Motherson\generated\" + FileName, FileMode.Create))
            {
                wb.Write(file);
                file.Close();
                Console.WriteLine("File Creation Successful");
            }

            #endregion

            return Ok();
        }

        public void WriteBodyContent(List<RuleHeader> Rule, XSSFSheet sheet, ICellStyle Checklist, ICellStyle Data)
        {
            int SR_NO = 1;
            int bodyRowIndex = 5;
            for (int i = 0; i < Rule.Count; i++)
            {
                IRow RuleHeader = sheet.CreateRow(bodyRowIndex);
                ICell ruleHeaderCell = RuleHeader.CreateCell(1);
                sheet.AddMergedRegion(new CellRangeAddress(bodyRowIndex, bodyRowIndex, 1, 6));
                for (int k = 1; k < 7; k++)
                {
                    var stylingCell = RuleHeader.CreateCell(k);
                    stylingCell.CellStyle = Checklist;
                }
                ruleHeaderCell.SetCellValue(Rule[i].RuleName);

                for (int j = 0; j < Rule[i].RuleListDetails.Count; j++)
                {
                    bodyRowIndex = bodyRowIndex + 1;
                    var detailRow = sheet.CreateRow(bodyRowIndex);

                    var sectionCell = detailRow.CreateCell(1);
                    sectionCell.CellStyle = Data;
                    sectionCell.SetCellValue(Rule[i].RuleListDetails[j].Section);

                    var typeCell = detailRow.CreateCell(2);
                    typeCell.CellStyle = Data;
                    typeCell.SetCellValue(Rule[i].RuleListDetails[j].Type);

                    var serialNoCell = detailRow.CreateCell(3);
                    serialNoCell.CellStyle = Data;
                    serialNoCell.SetCellValue(SR_NO++);

                    var checkpointCell = detailRow.CreateCell(4);
                    checkpointCell.CellStyle = Data;
                    checkpointCell.SetCellValue(Rule[i].RuleListDetails[j].checkpointStatement);

                    var scoreCell = detailRow.CreateCell(5);
                    scoreCell.CellStyle = Data;
                    scoreCell.SetCellValue(Rule[i].RuleListDetails[j].Score);

                    var remarkCell = detailRow.CreateCell(6);
                    remarkCell.CellStyle = Data;
                    remarkCell.SetCellValue(Rule[i].RuleListDetails[j].Remark);
                }

                bodyRowIndex += 1;
            }
        }
    }
}
