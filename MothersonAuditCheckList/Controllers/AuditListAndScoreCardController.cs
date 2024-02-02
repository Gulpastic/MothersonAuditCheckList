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


namespace MothersonAuditCheckList.Controllers
{
    public class AuditListAndScoreCardController : ControllerBase
    {
        [HttpPost]
        [ProducesResponseType(StatusCodes.Status200OK)]
        [ProducesResponseType(StatusCodes.Status400BadRequest)]
        [ProducesResponseType(StatusCodes.Status500InternalServerError)]
        [Route("api/CreateExcel")]
        public ActionResult createExcel([FromBody] AuditListDTO auditListDTO)
        {
            #region Create excel workbook and Sheets

            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet checklist = (XSSFSheet)wb.CreateSheet("Safety Audit Checklist");
            XSSFSheet card = (XSSFSheet)wb.CreateSheet("Score Card");

            #endregion

            #region Create Header Cell Style and Font

            var Header = wb.CreateCellStyle();
            Header.Alignment = HorizontalAlignment.Center;
            IFont HeaderFont = wb.CreateFont();
            HeaderFont.Boldweight = (short)FontBoldWeight.Bold;
            HeaderFont.FontName = "Calibri";
            HeaderFont.FontHeightInPoints = (short)12;
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
            checkpointHeaderFont.FontHeightInPoints = (short)12;
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
            dataFont.FontHeightInPoints = (short)11;
            Data.SetFont(dataFont);
            Data.BorderLeft = BorderStyle.Medium;
            Data.BorderRight = BorderStyle.Medium;
            Data.BorderBottom = BorderStyle.Medium;
            Data.BorderTop = BorderStyle.Medium;
            Data.WrapText = true;

            #endregion

            #region Create and Merge Cells

            IRow headerRow = checklist.CreateRow(1);
            var headerCell = headerRow.CreateCell(1);
            XSSFRichTextString richString = new XSSFRichTextString("COSA Safety Audit Checksheet-" + DateTime.Now.ToString("yyyy") + Environment.NewLine + "COMPREHENSIVE CHECKLIST OF ENVIRONMENT, HEALTH & SAFETY AUDIT");
            headerCell.SetCellValue(richString);

            checklist.AddMergedRegion(new CellRangeAddress(1, 1, 1, 6));
            checklist.AddMergedRegion(new CellRangeAddress(2, 2, 2, 3));
            checklist.AddMergedRegion(new CellRangeAddress(3, 3, 2, 3));
            headerCell.CellStyle = Header;
            for (int i = 1; i < 7; i++)
            {
                headerCell = headerRow.CreateCell(i);
                headerCell.CellStyle = Header;

            }

            #endregion

            #region add logo on top
            byte[] data = System.IO.File.ReadAllBytes("logo.png");
            int pictureIndex = wb.AddPicture(data, PictureType.PNG);
            ICreationHelper helper = wb.GetCreationHelper();
            IDrawing drawing = checklist.CreateDrawingPatriarch();
            IClientAnchor anchor = helper.CreateClientAnchor();
            anchor.Col1 = 1;
            anchor.Row1 = 1;
            anchor.Col2 = 1;
            anchor.Row2 = 1;
            IPicture picture = drawing.CreatePicture(anchor, pictureIndex);
            picture.Resize(0.8);
            anchor.AnchorType = AnchorType.MoveAndResize;
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
            headerCell2.CellStyle = Header;

            cellIndex = cellIndex + 1;
            headerCell2 = headerRow2.CreateCell(cellIndex);
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


            for (int i = 1; i < 15; i++)
                checklist.AutoSizeRow(i);

            for (int j = 1; j < 7; j++)
                checklist.AutoSizeColumn(j);

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
    }

}
