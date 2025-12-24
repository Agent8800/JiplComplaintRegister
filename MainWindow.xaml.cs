using System.Text.RegularExpressions;
using System.Windows;
using Microsoft.Win32;
using JiplComplaintRegister.Data;
using JiplComplaintRegister.Models;
using JiplComplaintRegister.Services;

namespace JiplComplaintRegister;

public partial class MainWindow : Window
{
    private readonly ComplaintRepository _repo;
    private List<Complaint> _currentList = new();
    private List<Complaint> _currentReport = new();

    public MainWindow()
    {
        InitializeComponent();

        var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
        var dbPath = System.IO.Path.Combine(appData, "JiplComplaintRegister", "complaints.db");
        _repo = new ComplaintRepository(dbPath);

        FromDate.SelectedDate = DateTime.Today.AddMonths(-1);
        ToDate.SelectedDate = DateTime.Today;

        ReportMonth.SelectedIndex = DateTime.Today.Month - 1;
        ReportYear.Text = DateTime.Today.Year.ToString();

        RefreshList();
        GenerateReport();
    }

    private void SaveComplaint_Click(object sender, RoutedEventArgs e)
    {
        if (!ValidateNew(out var c)) return;

        var no = _repo.Create(c);
        MessageBox.Show($"Complaint saved.\n\nComplaint No:\n{no}", "Saved", MessageBoxButton.OK, MessageBoxImage.Information);

        NameBox.Text = "";
        MobileBox.Text = "";
        LocationBox.Text = "";
        DepartmentBox.Text = "";
        ProductBox.Text = "";
        SerialBox.Text = "";
        DetailsBox.Text = "";

        RefreshList();
        GenerateReport();
    }

    private bool ValidateNew(out Complaint complaint)
    {
        complaint = new Complaint
        {
            Name = NameBox.Text.Trim(),
            Mobile = MobileBox.Text.Trim(),
            Location = LocationBox.Text.Trim(),
            Department = DepartmentBox.Text.Trim(),
            Product = ProductBox.Text.Trim(),
            SerialNo = SerialBox.Text.Trim(),
            Details = DetailsBox.Text.Trim()
        };

        if (string.IsNullOrWhiteSpace(complaint.Name)) { Msg("Name is required."); return false; }
        if (string.IsNullOrWhiteSpace(complaint.Mobile)) { Msg("Mobile is required."); return false; }
        if (!Regex.IsMatch(complaint.Mobile, @"^[0-9+\-\s]{6,20}$")) { Msg("Mobile format looks invalid."); return false; }
        if (string.IsNullOrWhiteSpace(complaint.Location)) { Msg("Location is required."); return false; }
        if (string.IsNullOrWhiteSpace(complaint.Department)) { Msg("Department is required."); return false; }
        if (string.IsNullOrWhiteSpace(complaint.Product)) { Msg("Product is required."); return false; }
        if (string.IsNullOrWhiteSpace(complaint.SerialNo)) { Msg("Serial Number is required."); return false; }

        return true;
    }

    private static void Msg(string text) =>
        MessageBox.Show(text, "Validation", MessageBoxButton.OK, MessageBoxImage.Warning);

    private void ApplyFilters_Click(object sender, RoutedEventArgs e) => RefreshList();

    private void RefreshList()
    {
        var status = ((System.Windows.Controls.ComboBoxItem)StatusFilter.SelectedItem).Content!.ToString()!;
        var from = FromDate.SelectedDate;
        var to = ToDate.SelectedDate;
        var search = SearchBox.Text ?? "";

        _currentList = _repo.List(status, from, to, search);
        ComplaintsGrid.ItemsSource = _currentList;
    }

    private Complaint? SelectedComplaint() => ComplaintsGrid.SelectedItem as Complaint;

    private void Edit_Click(object sender, RoutedEventArgs e)
    {
        var selected = SelectedComplaint();
        if (selected == null) { MessageBox.Show("Select a complaint row first."); return; }

        // Simple edit using an input dialog-style window would be nicer.
        // For minimal code: toggle status only is provided separately.
        selected.Details = selected.Details; // keep
        var name = Microsoft.VisualBasic.Interaction.InputBox("Edit Name:", "Edit", selected.Name);
        if (!string.IsNullOrWhiteSpace(name))
            selected.Name = name.Trim();

        _repo.Update(selected);
        RefreshList();
        GenerateReport();
    }

    private void Toggle_Click(object sender, RoutedEventArgs e)
    {
        var selected = SelectedComplaint();
        if (selected == null) { MessageBox.Show("Select a complaint row first."); return; }

        selected.Status = selected.Status == ComplaintRepository.Pending
            ? ComplaintRepository.Completed
            : ComplaintRepository.Pending;

        if (selected.Status == ComplaintRepository.Completed)
            selected.CompletedAt = DateTime.Now;
        else
            selected.CompletedAt = null;

        _repo.Update(selected);
        RefreshList();
        GenerateReport();
    }

    private void ExportExcel_Click(object sender, RoutedEventArgs e)
    {
        var dlg = new SaveFileDialog { Filter = "Excel (*.xlsx)|*.xlsx", FileName = "complaints_export.xlsx" };
        if (dlg.ShowDialog() != true) return;
        ExportService.ExportExcel(dlg.FileName, _currentList);
        MessageBox.Show("Excel exported.");
    }

    private void ExportPdf_Click(object sender, RoutedEventArgs e)
    {
        var dlg = new SaveFileDialog { Filter = "PDF (*.pdf)|*.pdf", FileName = "complaints_export.pdf" };
        if (dlg.ShowDialog() != true) return;
        ExportService.ExportPdf(dlg.FileName, "Complaints Export", _currentList);
        MessageBox.Show("PDF exported.");
    }

    private void GenerateReport_Click(object sender, RoutedEventArgs e) => GenerateReport();

    private void GenerateReport()
    {
        if (!int.TryParse(ReportYear.Text, out var year)) year = DateTime.Today.Year;
        var month = ReportMonth.SelectedIndex + 1;
        var status = ((System.Windows.Controls.ComboBoxItem)ReportStatus.SelectedItem).Content!.ToString()!;

        var (pending, completed, items) = _repo.MonthlyReport(year, month, status);
        _currentReport = items;

        ReportSummary.Text = $"Summary {year}-{month:00}: Pending={pending}  Completed={completed}  Total={items.Count}";
        ReportGrid.ItemsSource = _currentReport;
    }

    private void ReportExportExcel_Click(object sender, RoutedEventArgs e)
    {
        var dlg = new SaveFileDialog { Filter = "Excel (*.xlsx)|*.xlsx", FileName = "monthly_report.xlsx" };
        if (dlg.ShowDialog() != true) return;
        ExportService.ExportExcel(dlg.FileName, _currentReport);
        MessageBox.Show("Excel exported.");
    }

    private void ReportExportPdf_Click(object sender, RoutedEventArgs e)
    {
        var dlg = new SaveFileDialog { Filter = "PDF (*.pdf)|*.pdf", FileName = "monthly_report.pdf" };
        if (dlg.ShowDialog() != true) return;
        ExportService.ExportPdf(dlg.FileName, "Monthly Report", _currentReport);
        MessageBox.Show("PDF exported.");
    }
}
