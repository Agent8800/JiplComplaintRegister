using Microsoft.Data.Sqlite;
using JiplComplaintRegister.Models;
using System.Text.RegularExpressions;
using System.IO;

namespace JiplComplaintRegister.Data;

public class ComplaintRepository
{
    public const string Pending = "Pending";
    public const string Completed = "Completed";

    private readonly string _dbPath;

    public ComplaintRepository(string dbPath)
    {
        _dbPath = dbPath;
        Directory.CreateDirectory(Path.GetDirectoryName(_dbPath)!);
        using var con = Open();
        Init(con);
    }

    private SqliteConnection Open()
    {
        var con = new SqliteConnection($"Data Source={_dbPath}");
        con.Open();
        return con;
    }

    private static void Init(SqliteConnection con)
    {
        var cmd = con.CreateCommand();
        cmd.CommandText = """
        CREATE TABLE IF NOT EXISTS complaints (
            Id INTEGER PRIMARY KEY AUTOINCREMENT,
            ComplaintNo TEXT UNIQUE NOT NULL,
            CreatedAt TEXT NOT NULL,
            Name TEXT NOT NULL,
            Mobile TEXT NOT NULL,
            Location TEXT NOT NULL,
            Department TEXT NOT NULL,
            Product TEXT NOT NULL,
            SerialNo TEXT NOT NULL,
            Status TEXT NOT NULL,
            CompletedAt TEXT NULL,
            Details TEXT NOT NULL DEFAULT ''
        );

        CREATE INDEX IF NOT EXISTS idx_created ON complaints(CreatedAt);
        CREATE INDEX IF NOT EXISTS idx_status ON complaints(Status);
        """;
        cmd.ExecuteNonQuery();
    }

    private static string SanitizeToken(string s)
    {
        s = (s ?? "").Trim().ToUpperInvariant();
        s = s.Replace("/", "-");
        s = Regex.Replace(s, @"\s+", "-");
        s = Regex.Replace(s, @"[^A-Z0-9\-]", "");
        return string.IsNullOrWhiteSpace(s) ? "NA" : s;
    }

    private static string YyyyMmDd(DateTime d) => d.ToString("yyyyMMdd");

    private int NextSequence(SqliteConnection con, string locToken, string deptToken, string day)
    {
        var prefix = $"JIPL/{locToken}/{day}/{deptToken}/";

        using var cmd = con.CreateCommand();
        cmd.CommandText = """
        SELECT ComplaintNo
        FROM complaints
        WHERE ComplaintNo LIKE $p || '%'
        ORDER BY ComplaintNo DESC
        LIMIT 1
        """;
        cmd.Parameters.AddWithValue("$p", prefix);

        var last = cmd.ExecuteScalar() as string;
        if (string.IsNullOrWhiteSpace(last)) return 1;

        var parts = last.Split('/');
        return (parts.Length > 0 && int.TryParse(parts[^1], out var n)) ? n + 1 : 1;
    }

    public string Create(Complaint c)
    {
        var now = DateTime.Now;
        var day = YyyyMmDd(now);
        var locToken = SanitizeToken(c.Location);
        var deptToken = SanitizeToken(c.Department);

        using var con = Open();
        var seq = NextSequence(con, locToken, deptToken, day);
        var complaintNo = $"JIPL/{locToken}/{day}/{deptToken}/{seq:0000}";

        using var cmd = con.CreateCommand();
        cmd.CommandText = """
        INSERT INTO complaints
        (ComplaintNo, CreatedAt, Name, Mobile, Location, Department, Product, SerialNo, Status, CompletedAt, Details)
        VALUES
        ($no, $created, $name, $mobile, $loc, $dept, $prod, $serial, $status, $completed, $details)
        """;

        cmd.Parameters.AddWithValue("$no", complaintNo);
        cmd.Parameters.AddWithValue("$created", now.ToString("yyyy-MM-dd HH:mm:ss"));
        cmd.Parameters.AddWithValue("$name", c.Name.Trim());
        cmd.Parameters.AddWithValue("$mobile", c.Mobile.Trim());
        cmd.Parameters.AddWithValue("$loc", c.Location.Trim());
        cmd.Parameters.AddWithValue("$dept", c.Department.Trim());
        cmd.Parameters.AddWithValue("$prod", c.Product.Trim());
        cmd.Parameters.AddWithValue("$serial", c.SerialNo.Trim());
        cmd.Parameters.AddWithValue("$status", Pending);
        cmd.Parameters.AddWithValue("$completed", DBNull.Value);
        cmd.Parameters.AddWithValue("$details", (c.Details ?? "").Trim());

        cmd.ExecuteNonQuery();
        return complaintNo;
    }

    public List<Complaint> List(string status = "All", DateTime? from = null, DateTime? to = null, string search = "")
    {
        using var con = Open();
        using var cmd = con.CreateCommand();

        var where = new List<string>();
        if (status is Pending or Completed)
        {
            where.Add("Status = $status");
            cmd.Parameters.AddWithValue("$status", status);
        }
        if (from.HasValue)
        {
            where.Add("date(CreatedAt) >= date($from)");
            cmd.Parameters.AddWithValue("$from", from.Value.ToString("yyyy-MM-dd"));
        }
        if (to.HasValue)
        {
            where.Add("date(CreatedAt) <= date($to)");
            cmd.Parameters.AddWithValue("$to", to.Value.ToString("yyyy-MM-dd"));
        }
        if (!string.IsNullOrWhiteSpace(search))
        {
            where.Add("""
            (ComplaintNo LIKE $s OR Name LIKE $s OR Mobile LIKE $s OR Location LIKE $s OR Department LIKE $s OR Product LIKE $s OR SerialNo LIKE $s)
            """);
            cmd.Parameters.AddWithValue("$s", $"%{search.Trim()}%");
        }

        cmd.CommandText = $"""
        SELECT Id, ComplaintNo, CreatedAt, Name, Mobile, Location, Department, Product, SerialNo, Status, CompletedAt, Details
        FROM complaints
        {(where.Count > 0 ? "WHERE " + string.Join(" AND ", where) : "")}
        ORDER BY datetime(CreatedAt) DESC
        """;

        using var r = cmd.ExecuteReader();
        var list = new List<Complaint>();
        while (r.Read())
        {
            list.Add(new Complaint
            {
                Id = r.GetInt64(0),
                ComplaintNo = r.GetString(1),
                CreatedAt = DateTime.Parse(r.GetString(2)),
                Name = r.GetString(3),
                Mobile = r.GetString(4),
                Location = r.GetString(5),
                Department = r.GetString(6),
                Product = r.GetString(7),
                SerialNo = r.GetString(8),
                Status = r.GetString(9),
                CompletedAt = r.IsDBNull(10) ? null : DateTime.Parse(r.GetString(10)),
                Details = r.IsDBNull(11) ? "" : r.GetString(11),
            });
        }
        return list;
    }

    public void Update(Complaint c)
    {
        using var con = Open();
        using var cmd = con.CreateCommand();

        DateTime? completedAt = c.Status == Completed ? (c.CompletedAt ?? DateTime.Now) : null;

        cmd.CommandText = """
        UPDATE complaints SET
            Name=$name, Mobile=$mobile, Location=$loc, Department=$dept, Product=$prod, SerialNo=$serial,
            Status=$status, CompletedAt=$completed, Details=$details
        WHERE ComplaintNo=$no
        """;

        cmd.Parameters.AddWithValue("$no", c.ComplaintNo);
        cmd.Parameters.AddWithValue("$name", c.Name.Trim());
        cmd.Parameters.AddWithValue("$mobile", c.Mobile.Trim());
        cmd.Parameters.AddWithValue("$loc", c.Location.Trim());
        cmd.Parameters.AddWithValue("$dept", c.Department.Trim());
        cmd.Parameters.AddWithValue("$prod", c.Product.Trim());
        cmd.Parameters.AddWithValue("$serial", c.SerialNo.Trim());
        cmd.Parameters.AddWithValue("$status", c.Status);
        cmd.Parameters.AddWithValue("$completed", completedAt == null ? DBNull.Value : completedAt.Value.ToString("yyyy-MM-dd HH:mm:ss"));
        cmd.Parameters.AddWithValue("$details", (c.Details ?? "").Trim());

        cmd.ExecuteNonQuery();
    }

    public (int pending, int completed, List<Complaint> items) MonthlyReport(int year, int month, string status = "All")
    {
        var from = new DateTime(year, month, 1);
        var to = from.AddMonths(1); // exclusive end

        using var con = Open();
        using var cmd = con.CreateCommand();

        cmd.CommandText = $"""
        SELECT Id, ComplaintNo, CreatedAt, Name, Mobile, Location, Department, Product, SerialNo, Status, CompletedAt, Details
        FROM complaints
        WHERE date(CreatedAt) >= date($from)
          AND date(CreatedAt) <  date($to)
          {(status is Pending or Completed ? "AND Status=$status" : "")}
        ORDER BY datetime(CreatedAt) DESC
        """;

        cmd.Parameters.AddWithValue("$from", from.ToString("yyyy-MM-dd"));
        cmd.Parameters.AddWithValue("$to", to.ToString("yyyy-MM-dd"));
        if (status is Pending or Completed)
            cmd.Parameters.AddWithValue("$status", status);

        using var r = cmd.ExecuteReader();
        var list = new List<Complaint>();
        while (r.Read())
        {
            list.Add(new Complaint
            {
                Id = r.GetInt64(0),
                ComplaintNo = r.GetString(1),
                CreatedAt = DateTime.Parse(r.GetString(2)),
                Name = r.GetString(3),
                Mobile = r.GetString(4),
                Location = r.GetString(5),
                Department = r.GetString(6),
                Product = r.GetString(7),
                SerialNo = r.GetString(8),
                Status = r.GetString(9),
                CompletedAt = r.IsDBNull(10) ? null : DateTime.Parse(r.GetString(10)),
                Details = r.IsDBNull(11) ? "" : r.GetString(11),
            });
        }

        var p = list.Count(x => x.Status == Pending);
        var c = list.Count(x => x.Status == Completed);
        return (p, c, list);
    }
}
