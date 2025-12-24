namespace JiplComplaintRegister.Models;

public class Complaint
{
    public long Id { get; set; }
    public string ComplaintNo { get; set; } = "";
    public DateTime CreatedAt { get; set; }
    public string Name { get; set; } = "";
    public string Mobile { get; set; } = "";
    public string Location { get; set; } = "";
    public string Department { get; set; } = "";
    public string Product { get; set; } = "";
    public string SerialNo { get; set; } = "";
    public string Status { get; set; } = "Pending"; // Pending | Completed
    public DateTime? CompletedAt { get; set; }
    public string Details { get; set; } = "";
}
