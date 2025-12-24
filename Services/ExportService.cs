using ClosedXML.Excel;
using JiplComplaintRegister.Models;
using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;

namespace JiplComplaintRegister.Services;

public static class ExportService
{
    private static readonly string[] Headers =
    [
        "Complaint No","Created At","Name","Mobile","Location","Department","Product","Serial No","Status","Completed At"
    ];

    public static void ExportExcel(string path, IEnumerable<Complaint> items)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Complaints");

        for (int i = 0; i < Headers.Length; i++)
            ws.Cell(1, i + 1).Value = Headers[i];

        int row = 2;
        foreach (var c in items)
        {
            ws.Cell(row, 1).Value = c.ComplaintNo;
            ws.Cell(row, 2).Value = c.CreatedAt.ToString("yyyy-MM-dd HH:mm:ss");
            ws.Cell(row, 3).Value = c.Name;
            ws.Cell(row, 4).Value = c.Mobile;
            ws.Cell(row, 5).Value = c.Location;
            ws.Cell(row, 6).Value = c.Department;
            ws.Cell(row, 7).Value = c.Product;
            ws.Cell(row, 8).Value = c.SerialNo;
            ws.Cell(row, 9).Value = c.Status;
            ws.Cell(row,10).Value = c.CompletedAt?.ToString("yyyy-MM-dd HH:mm:ss") ?? "";
            row++;
        }

        ws.Range(1, 1, 1, Headers.Length).Style.Font.Bold = true;
        ws.Columns().AdjustToContents();
        wb.SaveAs(path);
    }

    public static void ExportPdf(string path, string title, IEnumerable<Complaint> items)
    {
        QuestPDF.Settings.License = LicenseType.Community;

        Document.Create(container =>
        {
            container.Page(page =>
            {
                page.Size(PageSizes.A4.Landscape());
                page.Margin(20);
                page.DefaultTextStyle(x => x.FontSize(9));

                page.Header().Column(col =>
                {
                    col.Item().Text(title).FontSize(16).SemiBold();
                    col.Item().Text($"Generated: {DateTime.Now:yyyy-MM-dd HH:mm:ss}").FontColor(Colors.Grey.Darken2);
                });

                page.Content().Table(table =>
                {
                    table.ColumnsDefinition(cols =>
                    {
                        cols.RelativeColumn(2);
                        cols.RelativeColumn(2);
                        cols.RelativeColumn(2);
                        cols.RelativeColumn(2);
                        cols.RelativeColumn(2);
                        cols.RelativeColumn(2);
                        cols.RelativeColumn(2);
                        cols.RelativeColumn(2);
                        cols.RelativeColumn(1);
                        cols.RelativeColumn(2);
                    });

                    table.Header(h =>
                    {
                        foreach (var head in Headers)
                            h.Cell().Background("#EEF2FF").Padding(4).Text(head).SemiBold();
                    });

                    foreach (var c in items)
                    {
                        table.Cell().Padding(3).Text(c.ComplaintNo);
                        table.Cell().Padding(3).Text(c.CreatedAt.ToString("yyyy-MM-dd HH:mm:ss"));
                        table.Cell().Padding(3).Text(c.Name);
                        table.Cell().Padding(3).Text(c.Mobile);
                        table.Cell().Padding(3).Text(c.Location);
                        table.Cell().Padding(3).Text(c.Department);
                        table.Cell().Padding(3).Text(c.Product);
                        table.Cell().Padding(3).Text(c.SerialNo);
                        table.Cell().Padding(3).Text(c.Status);
                        table.Cell().Padding(3).Text(c.CompletedAt?.ToString("yyyy-MM-dd HH:mm:ss") ?? "");
                    }
                });

                page.Footer().AlignRight().Text(x =>
                {
                    x.Span("Page ");
                    x.CurrentPageNumber();
                    x.Span(" / ");
                    x.TotalPages();
                });
            });
        }).GeneratePdf(path);
    }
}
