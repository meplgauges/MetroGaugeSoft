using DocumentFormat.OpenXml.EMMA;
using Microsoft.Data.SqlClient;
using System;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.Windows;


namespace EVMS.Service
{
    public class DataStorageService : IDisposable
    {
        private readonly string _connectionString;

        public DataStorageService()
        {
            _connectionString = ConfigurationManager.ConnectionStrings["EVMSDb"].ConnectionString;

            if (string.IsNullOrWhiteSpace(_connectionString))
                throw new ArgumentException("Connection string must not be empty.", nameof(_connectionString));
        }

        // Get PartConfig list by part number
        public List<PartReadingDataModel> GetPartConfigByPartNumber(string partNumber)
        {
            var list = new List<PartReadingDataModel>();
            string query = "SELECT * FROM PartConfig WHERE Para_No = @PartNumber AND IsEnabled = 1";

            using SqlConnection conn = new(_connectionString);
            using SqlCommand cmd = new(query, conn);
            cmd.Parameters.AddWithValue("@PartNumber", partNumber);

            conn.Open();
            using SqlDataReader reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                list.Add(new PartReadingDataModel
                {
                    Para_No = reader["Para_No"].ToString(),
                    Parameter = reader["Parameter"].ToString(),
                    ShortName = reader["ShortName"].ToString(),
                    D_Name= reader["D_Name"].ToString(),
                    Nominal = Convert.ToDouble(reader["Nominal"]),
                    RTolPlus = Convert.ToDouble(reader["RTolPlus"]),
                    RTolMinus = Convert.ToDouble(reader["RTolMinus"]),
                    Sign_Change=Convert.ToInt32(reader["Sign_Change"]),
                    Compensation = Convert.ToDouble(reader["Compensation"])

                });
            }
            return list;
        }

        // Update Sign_Change column
        public bool UpdateSignChange(string paraNo, string Parameter, int signChangeValue)
        {
            try
            {
                string query = "UPDATE PartConfig SET Sign_Change = @SignChange " +
                               "WHERE Para_No = @ParaNo AND Parameter = @Parameter";

                using SqlConnection conn = new(_connectionString);
                using SqlCommand cmd = new(query, conn);
                cmd.Parameters.AddWithValue("@SignChange", signChangeValue);
                cmd.Parameters.AddWithValue("@ParaNo", paraNo);
                cmd.Parameters.AddWithValue("@Parameter", Parameter);

                conn.Open();
                int rowsAffected = cmd.ExecuteNonQuery();
                return rowsAffected > 0;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error updating Sign_Change: {ex.Message}");
                return false;
            }
        }

        // Update Compensation column
        public bool UpdateCompensation(string paraNo, string Parameter, double compensationValue)
        {
            try
            {
                string query = "UPDATE PartConfig SET Compensation = @Compensation " +
                               "WHERE Para_No = @ParaNo AND Parameter = @Parameter";

                using SqlConnection conn = new(_connectionString);
                using SqlCommand cmd = new(query, conn);
                cmd.Parameters.AddWithValue("@Compensation", compensationValue);
                cmd.Parameters.AddWithValue("@ParaNo", paraNo);
                cmd.Parameters.AddWithValue("@Parameter", Parameter);

                conn.Open();
                int rowsAffected = cmd.ExecuteNonQuery();
                return rowsAffected > 0;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error updating Compensation: {ex.Message}");
                return false;
            }
        }

        // Get ProbeInstallationData by part number
        public List<ProbeInstallModel> GetProbeInstallByPartNumber(string partNumber)
        {
            var list = new List<ProbeInstallModel>();
            string query = @"
                SELECT PartNo, ProbeId, Name
                FROM ProbeInstallationData
                WHERE PartNo = @PartNo";

            using SqlConnection conn = new(_connectionString);
            using SqlCommand cmd = new(query, conn);
            cmd.Parameters.AddWithValue("@PartNo", partNumber);

            conn.Open();
            using SqlDataReader reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                var item = new ProbeInstallModel
                {
                    PartNo = reader["PartNo"].ToString(),
                    ProbeId = reader["ProbeId"].ToString(),
                    Name = reader["Name"].ToString()
                };
                list.Add(item);
            }
            return list;
        }



        public List<PartConfigModel> GetPartConfig(string partNumber)
        {
            var list = new List<PartConfigModel>();
            string query = @"
                                SELECT Parameter, Nominal, RTolPlus, RTolMinus, YTolPlus, YTolMinus, ProbeStatus, Para_No, IsEnabled,Sign_Change,Compensation
                                FROM PartConfig
                                WHERE Para_No = @ParaNo AND IsEnabled = 1";


            using SqlConnection conn = new(_connectionString);
            using SqlCommand cmd = new(query, conn);
            cmd.Parameters.AddWithValue("@ParaNo", partNumber);

            conn.Open();
            using SqlDataReader reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                list.Add(new PartConfigModel
                {
                    Parameter = reader["Parameter"].ToString(),
                    Nominal = Convert.ToDouble(reader["Nominal"]),
                    RTolPlus = Convert.ToDouble(reader["RTolPlus"]),
                    RTolMinus = Convert.ToDouble(reader["RTolMinus"]),
                    YTolPlus = Convert.ToDouble(reader["YTolPlus"]),
                    YTolMinus = Convert.ToDouble(reader["YTolMinus"]),
                    Sign_Change = Convert.ToInt32(reader["Sign_Change"]),
                    Compensation = Convert.ToDouble(reader["Compensation"]),
                    Para_No = reader["Para_No"].ToString()


                });
            }
            return list;
        }



        // Get MasterReadingData by part number
        public List<MasterReadingModel> GetMasterReadingByPart(string partNumber)
        {
            var list = new List<MasterReadingModel>();

            string query = @"
        SELECT 
            p.SrNo,
            p.Parameter,
            COALESCE(m.Nominal, p.Nominal) AS Nominal,
            COALESCE(m.RTolPlus, p.RTolPlus) AS RTolPlus,
            COALESCE(m.RTolMinus, p.RTolMinus) AS RTolMinus
        FROM PartConfig p
        LEFT JOIN MasterReadingData m
            ON p.Parameter = m.Parameter      -- match by parameter name
           AND m.Para_No = @PartNumber        -- use master data for this part
        WHERE p.Para_No = @PartNumber         -- only parameters for this part
        ORDER BY p.SrNo;                      -- preserve PartConfig order";

            using SqlConnection conn = new(_connectionString);
            using SqlCommand cmd = new(query, conn);
            cmd.Parameters.AddWithValue("@PartNumber", partNumber);

            conn.Open();
            using SqlDataReader reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                list.Add(new MasterReadingModel
                {
                    Para_No = reader["SrNo"].ToString(),
                    Parameter = reader["Parameter"].ToString(),
                    D_Name= reader["Parameter"].ToString(),
                    Nominal = Convert.ToDouble(reader["Nominal"]),
                    RTolPlus = Convert.ToDouble(reader["RTolPlus"]),
                    RTolMinus = Convert.ToDouble(reader["RTolMinus"])
                });
            }

            return list;
        }

        // Get Active parts from Part_Entry table (ActivePart = 1)
        public List<PartEntryModel> GetActiveParts()
        {
            var list = new List<PartEntryModel>();
            string query = "SELECT Para_No, Para_Name, ActivePart FROM Part_Entry WHERE ActivePart = 1";

            using SqlConnection conn = new(_connectionString);
            using SqlCommand cmd = new(query, conn);

            conn.Open();
            using SqlDataReader reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                list.Add(new PartEntryModel
                {
                    Para_No = reader["Para_No"].ToString(),
                    Para_Name = reader["Para_Name"].ToString(),
                    ActivePart = Convert.ToInt32(reader["ActivePart"])
                });
            }
            return list;
        }


        public bool ProbeReferencesExist(string partNo)
        {
            string query = "SELECT COUNT(*) FROM MasterReadingProbeReference WHERE PartNo = @PartNo";

            using SqlConnection conn = new(_connectionString);
            using SqlCommand cmd = new(query, conn);
            cmd.Parameters.AddWithValue("@PartNo", partNo);

            conn.Open();
            int count = (int)cmd.ExecuteScalar();
            return count > 0;
        }

        public List<(string Name, double Value)> GetMasterProbeRef(string partNo)
        {
            string query = @"
        SELECT Name, Value
        FROM MasterReadingProbeReference
        WHERE PartNo = @PartNo";

            var result = new List<(string, double)>();

            using SqlConnection conn = new SqlConnection(_connectionString);
            using SqlCommand cmd = new SqlCommand(query, conn);
            cmd.Parameters.AddWithValue("@PartNo", partNo);

            conn.Open();
            using SqlDataReader reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                string name = reader["Name"].ToString() ?? "";
                double value = reader["Value"] != DBNull.Value ? Convert.ToDouble(reader["Value"]) : 0.0;

                result.Add((name, value));
            }
            return result;
        }


        public void SaveProbeReadings(List<ProbeInstallModel> probes, string partNo, Dictionary<string, double> probeValues)
        {
            using SqlConnection conn = new SqlConnection(_connectionString);
            conn.Open();

            foreach (var probe in probes)
            {
                if (!string.IsNullOrEmpty(probe.ProbeId) && probeValues.TryGetValue(probe.ProbeId, out double value))
                {
                    string query = @"
                    MERGE MasterReadingProbeReference AS target
                    USING (VALUES (@PartNo, @ProbeId, @Name, @Value)) AS source (PartNo, ProbeId, Name, Value)
                    ON (target.PartNo = source.PartNo AND target.ProbeId = source.ProbeId)
                    WHEN MATCHED THEN
                        UPDATE SET Value = source.Value, LastUpdated = GETDATE(), Name = source.Name
                    WHEN NOT MATCHED THEN
                        INSERT (PartNo, ProbeId, Name, Value, LastUpdated)
                        VALUES (source.PartNo, source.ProbeId, source.Name, source.Value, GETDATE());
                ";

                    using SqlCommand cmd = new SqlCommand(query, conn);
                    cmd.Parameters.Clear();
                    cmd.Parameters.AddWithValue("@PartNo", partNo);
                    cmd.Parameters.AddWithValue("@ProbeId", probe.ProbeId);
                    cmd.Parameters.AddWithValue("@Name", probe.Name ?? "");
                    cmd.Parameters.AddWithValue("@Value", value);

                    cmd.ExecuteNonQuery();
                }
            }
        }

        public List<Controls> GetActiveBit()
        {
            var list = new List<Controls>();
            string query = "SELECT Description, Bit, code FROM Controls";

            using SqlConnection conn = new(_connectionString);
            using SqlCommand cmd = new(query, conn);

            conn.Open();
            using SqlDataReader reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                list.Add(new Controls
                {
                    Description = reader["Description"]?.ToString() ?? string.Empty,
                    Bit = reader["Bit"] != DBNull.Value ? Convert.ToInt32(reader["Bit"]) : 0,
                    Code = reader["Code"]?.ToString() ?? string.Empty,

                });
            }
            return list;
        }

        public List<Dictionary<string, object>> GetAllMeasurementReadingsDynamic(string partNo, DateTime? filterDate = null)
        {
            var list = new List<Dictionary<string, object>>();

            string query = "SELECT * FROM MeasurementReading WHERE PartNo = @PartNo";

            if (filterDate.HasValue)
            {
                query += " AND MeasurementDate >= @StartDate AND MeasurementDate < @EndDate ";
            }

            query += " ORDER BY MeasurementDate ASC";

            try
            {
                using var conn = new SqlConnection(_connectionString);
                using var cmd = new SqlCommand(query, conn);

                cmd.Parameters.AddWithValue("@PartNo", partNo);

                if (filterDate.HasValue)
                {
                    var startDate = filterDate.Value.Date;
                    var endDate = startDate.AddDays(1);

                    cmd.Parameters.AddWithValue("@StartDate", startDate);
                    cmd.Parameters.AddWithValue("@EndDate", endDate);
                }

                conn.Open();

                using var reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    var row = new Dictionary<string, object>();
                    for (int i = 1; i < reader.FieldCount; i++)
                    {
                        var colName = reader.GetName(i);
                        var val = reader.IsDBNull(i) ? null : reader.GetValue(i);
                        row[colName] = val;
                    }
                    list.Add(row);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("DB Exception: " + ex.Message);
            }

            return list;
        }


        public void UpdateAutoManualBit(int bitValue)
        {
            string sql = "UPDATE Controls SET Bit = @bit WHERE Description = 'Auto/Manual'";

            using (var conn = new SqlConnection(_connectionString))
            using (var cmd = new SqlCommand(sql, conn))
            {
                cmd.Parameters.AddWithValue("@bit", bitValue);
                conn.Open();
                cmd.ExecuteNonQuery();
            }
        }

        public int GetAutoManualBit()
        {
            string sql = "SELECT Bit FROM Controls WHERE Description = 'Auto/Manual'";

            using (var conn = new SqlConnection(_connectionString))
            using (var cmd = new SqlCommand(sql, conn))
            {
                conn.Open();
                var result = cmd.ExecuteScalar();
                return result != null ? Convert.ToInt32(result) : 0; // Default to Manual (0)
            }
        }



        // Insert a new record into MasterInspection table
        public async Task InsertMasterInspectionAsync(string partNo, string operatorId, string lotNo,
    float ol, float de, float hd, float gp, float stdg, float stdu, float girDia,
    float stn, float ovalitySdg, float ovalitySdu, float ovalityHead,
    float stemTaper, float efro, float faceRunout, float sh, float sRo, float dg,
    string status)
        {
            string query = @"
        INSERT INTO MasterInspection
        (PartNo, Operator_ID, LotNo, OL, DE, HD, GP, STDG, STDU, GIR_DIA, STN, Ovality_SDG, Ovality_SDU, Ovality_Head, Stem_Taper, EFRO, Face_Runout, SH, S_RO, DG, Status, InspectionDate)
        VALUES 
        (@PartNo, @Operator_ID, @LotNo, @OL, @DE, @HD, @GP, @STDG, @STDU, @GIR_DIA, @STN, @Ovality_SDG, @Ovality_SDU, @Ovality_Head, @Stem_Taper, @EFRO, @Face_Runout, @SH, @S_RO, @DG, @Status, GETDATE())";

            using var connection = new SqlConnection(_connectionString);
            using var command = new SqlCommand(query, connection);

            command.Parameters.AddWithValue("@PartNo", partNo);
            command.Parameters.AddWithValue("@Operator_ID", operatorId);
            command.Parameters.AddWithValue("@LotNo", lotNo);
            command.Parameters.AddWithValue("@OL", ol.ToString("F3"));
            command.Parameters.AddWithValue("@DE", de.ToString("F3"));
            command.Parameters.AddWithValue("@HD", hd.ToString("F3"));
            command.Parameters.AddWithValue("@GP", gp.ToString("F3"));
            command.Parameters.AddWithValue("@STDG", stdg.ToString("F3"));
            command.Parameters.AddWithValue("@STDU", stdu.ToString("F3"));
            command.Parameters.AddWithValue("@GIR_DIA", girDia.ToString("F3"));
            command.Parameters.AddWithValue("@STN", stn.ToString("F3"));
            command.Parameters.AddWithValue("@Ovality_SDG", ovalitySdg.ToString("F3"));
            command.Parameters.AddWithValue("@Ovality_SDU", ovalitySdu.ToString("F3"));
            command.Parameters.AddWithValue("@Ovality_Head", ovalityHead.ToString("F3"));
            command.Parameters.AddWithValue("@Stem_Taper", stemTaper.ToString("F3"));
            command.Parameters.AddWithValue("@EFRO", efro.ToString("F3"));
            command.Parameters.AddWithValue("@Face_Runout", faceRunout.ToString("F3"));
            command.Parameters.AddWithValue("@SH", sh.ToString("F3"));
            command.Parameters.AddWithValue("@S_RO", sRo.ToString("F3"));
            command.Parameters.AddWithValue("@DG", dg.ToString("F3"));
            command.Parameters.AddWithValue("@Status", status);

            await connection.OpenAsync();
            int rowsAffected = await command.ExecuteNonQueryAsync();
            if (rowsAffected == 0)
                throw new Exception("Insert failed: No rows were affected.");
        }


        public async Task InsertMeasurementReadingAsync(
    string partNo,
    string operatorId,
    string lotNo,
    float ol, float de, float hd, float gp, float stdg, float stdu, float girDia,
    float stn, float ovalitySdg, float ovalitySdu, float ovalityHead, float stemTaper,
    float efro, float faceRunout, float sh, float sRo, float dg,
    string status)
        {
            string query = @"
    INSERT INTO MeasurementReading
    (PartNo, Operator_ID, LotNo, OL, DE, HD, GP, STDG, STDU, GIR_DIA, STN, Ovality_SDG, Ovality_SDU, Ovality_Head, Stem_Taper, EFRO, Face_Runout, SH, S_RO, DG, Status, MeasurementDate)
    VALUES
    (@PartNo, @Operator_ID, @LotNo, @OL, @DE, @HD, @GP, @STDG, @STDU, @GIR_DIA, @STN, @Ovality_SDG, @Ovality_SDU, @Ovality_Head, @Stem_Taper, @EFRO, @Face_Runout, @SH, @S_RO, @DG, @Status, GETDATE())";

            using var connection = new SqlConnection(_connectionString);
            using var command = new SqlCommand(query, connection);

            command.Parameters.AddWithValue("@PartNo", partNo);
            command.Parameters.AddWithValue("@Operator_ID", operatorId);
            command.Parameters.AddWithValue("@LotNo", lotNo);
            command.Parameters.AddWithValue("@OL", ol.ToString("F3"));
            command.Parameters.AddWithValue("@DE", de.ToString("F3"));
            command.Parameters.AddWithValue("@HD", hd.ToString("F3"));
            command.Parameters.AddWithValue("@GP", gp.ToString("F3"));
            command.Parameters.AddWithValue("@STDG", stdg.ToString("F3"));
            command.Parameters.AddWithValue("@STDU", stdu.ToString("F3"));
            command.Parameters.AddWithValue("@GIR_DIA", girDia.ToString("F3"));
            command.Parameters.AddWithValue("@STN", stn.ToString("F3"));
            command.Parameters.AddWithValue("@Ovality_SDG", ovalitySdg.ToString("F3"));
            command.Parameters.AddWithValue("@Ovality_SDU", ovalitySdu.ToString("F3"));
            command.Parameters.AddWithValue("@Ovality_Head", ovalityHead.ToString("F3"));
            command.Parameters.AddWithValue("@Stem_Taper", stemTaper.ToString("F3"));
            command.Parameters.AddWithValue("@EFRO", efro.ToString("F3"));
            command.Parameters.AddWithValue("@Face_Runout", faceRunout.ToString("F3"));
            command.Parameters.AddWithValue("@SH", sh.ToString("F3"));
            command.Parameters.AddWithValue("@S_RO", sRo.ToString("F3"));
            command.Parameters.AddWithValue("@DG", dg.ToString("F3"));
            command.Parameters.AddWithValue("@Status", status);

            await connection.OpenAsync();
            int rowsAffected = await command.ExecuteNonQueryAsync();
            if (rowsAffected == 0)
                throw new Exception("Insert failed: No rows were affected.");
        }

        public async Task<DataTable> GetParameterDataTableAsync(string partNo)
        {
            DataTable dt = new DataTable();

            string query = @"
        SELECT Parameter, Value
        FROM MeasurementTable  -- Replace with your actual table
        WHERE PartNo = @PartNo";

            using (var conn = new SqlConnection(_connectionString))
            using (var cmd = new SqlCommand(query, conn))
            {
                cmd.Parameters.AddWithValue("@PartNo", partNo);
                await conn.OpenAsync();
                using (var reader = await cmd.ExecuteReaderAsync())
                {
                    dt.Load(reader);
                }
            }

            return dt;
        }



        public (int Mode, int SetValue, DateTime UpdatedAt) GetMasterExpiration()
        {
            using var con = new SqlConnection(_connectionString); // class member
            con.Open();
            using var cmd = new SqlCommand("SELECT Mode, SetValue, UpdatedAt FROM MasterExpiration WHERE Id = 1", con);
            using var reader = cmd.ExecuteReader();
            if (reader.Read())
            {
                bool modeBool = reader.GetBoolean(0); // Correct for BIT type
                int mode = modeBool ? 1 : 0;          // Convert to int
                return (mode, reader.GetInt32(1), reader.GetDateTime(2));
            }
            throw new InvalidOperationException("Master expiration settings not found.");
        }


        public async Task<InspectionData?> SelectInspectionDataAsync(string _model, string _lotNo, string _userId)
        {
            string sql = @"SELECT PartNo, LotNo, OperatorID, InspectionQty, OkCount 
                       FROM InspectionData
                       WHERE PartNo = @PartNo AND LotNo = @LotNo AND OperatorID = @OperatorID";

            using var con = new SqlConnection(_connectionString);
            await con.OpenAsync();

            using var cmd = new SqlCommand(sql, con);
            cmd.Parameters.AddWithValue("@PartNo", _model);
            cmd.Parameters.AddWithValue("@LotNo", _lotNo);
            cmd.Parameters.AddWithValue("@OperatorID", _userId);

            using var reader = await cmd.ExecuteReaderAsync();
            if (await reader.ReadAsync())
            {
                return new InspectionData
                {
                    PartNo = reader.GetString(0),
                    LotNo = reader.GetString(1),
                    OperatorID = reader.GetString(2),
                    InspectionQty = reader.GetInt32(3),
                    OkCount = reader.GetInt32(4)
                };
            }
            return null;
        }
        public async Task<List<string>> GetLotNumbersByPartNoAsync(string partNo)
        {
            var lotNumbers = new List<string>();
            string sql = "SELECT DISTINCT LotNo FROM InspectionData WHERE PartNo = @PartNo";

            using var con = new SqlConnection(_connectionString);
            await con.OpenAsync();

            using var cmd = new SqlCommand(sql, con);
            cmd.Parameters.AddWithValue("@PartNo", partNo);

            using var reader = await cmd.ExecuteReaderAsync();
            while (await reader.ReadAsync())
            {
                lotNumbers.Add(reader.GetString(0));
            }
            return lotNumbers;
        }
        public async Task<List<string>> GetOperatorsByPartNoAsync(string partNo)
        {
            var operators = new List<string>();
            string sql = "SELECT DISTINCT OperatorID FROM InspectionData WHERE PartNo = @PartNo";

            using var con = new SqlConnection(_connectionString);
            await con.OpenAsync();

            using var cmd = new SqlCommand(sql, con);
            cmd.Parameters.AddWithValue("@PartNo", partNo);

            using var reader = await cmd.ExecuteReaderAsync();
            while (await reader.ReadAsync())
            {
                operators.Add(reader.GetString(0));
            }
            return operators;
        }
        public async Task<List<string>> GetLotNumbersByPartAndDateRangeAsync(string partNo, DateTime? dateFrom, DateTime? dateTo)
        {
            var lotNumbers = new List<string>();

            using (SqlConnection con = new SqlConnection(_connectionString))
            {
                await con.OpenAsync();

                string query = @"SELECT DISTINCT LotNo FROM MeasurementReading WHERE 1=1";

                if (!string.IsNullOrEmpty(partNo))
                    query += " AND PartNo = @PartNo";

                if (dateFrom.HasValue)
                    query += " AND MeasurementDate >= @DateFrom";

                if (dateTo.HasValue)
                    query += " AND MeasurementDate <= @DateTo";

                using (SqlCommand cmd = new SqlCommand(query, con))
                {
                    if (!string.IsNullOrEmpty(partNo))
                        cmd.Parameters.AddWithValue("@PartNo", partNo);

                    if (dateFrom.HasValue)
                        cmd.Parameters.AddWithValue("@DateFrom", dateFrom.Value);

                    if (dateTo.HasValue)
                        cmd.Parameters.AddWithValue("@DateTo", dateTo.Value);

                    using (var reader = await cmd.ExecuteReaderAsync())
                    {
                        while (await reader.ReadAsync())
                        {
                            if (!reader.IsDBNull(reader.GetOrdinal("LotNo")))
                                lotNumbers.Add(reader.GetString(reader.GetOrdinal("LotNo")));
                        }
                    }
                }
            }

            return lotNumbers;
        }

        public async Task<List<string>> GetOperatorsByPartAndDateRangeAsync(string partNo, DateTime? dateFrom, DateTime? dateTo)
        {
            var operators = new List<string>();

            using (SqlConnection con = new SqlConnection(_connectionString))
            {
                await con.OpenAsync();

                string query = @"SELECT DISTINCT Operator_ID FROM MeasurementReading WHERE 1=1";

                if (!string.IsNullOrEmpty(partNo))
                    query += " AND PartNo = @PartNo";

                if (dateFrom.HasValue)
                    query += " AND MeasurementDate >= @DateFrom";

                if (dateTo.HasValue)
                    query += " AND MeasurementDate <= @DateTo";

                using (SqlCommand cmd = new SqlCommand(query, con))
                {
                    if (!string.IsNullOrEmpty(partNo))
                        cmd.Parameters.AddWithValue("@PartNo", partNo);

                    if (dateFrom.HasValue)
                        cmd.Parameters.AddWithValue("@DateFrom", dateFrom.Value);

                    if (dateTo.HasValue)
                        cmd.Parameters.AddWithValue("@DateTo", dateTo.Value);

                    using (var reader = await cmd.ExecuteReaderAsync())
                    {
                        while (await reader.ReadAsync())
                        {
                            if (!reader.IsDBNull(reader.GetOrdinal("Operator_ID")))
                                operators.Add(reader.GetString(reader.GetOrdinal("Operator_ID")));
                        }
                    }
                }
            }

            return operators;
        }

        public async Task InsertInspectionDataAsync(string _model, string _lotNo, string _userId)
        {
            string sql = @"INSERT INTO InspectionData (PartNo, LotNo, OperatorID, InspectionQty, OkCount)
                       VALUES (@PartNo, @LotNo, @OperatorID, 0, 0)";

            using var con = new SqlConnection(_connectionString);
            await con.OpenAsync();

            using var cmd = new SqlCommand(sql, con);
            cmd.Parameters.AddWithValue("@PartNo", _model);   // The name @PartNo in SQL must be exactly "@PartNo" here
            cmd.Parameters.AddWithValue("@LotNo", _lotNo);
            cmd.Parameters.AddWithValue("@OperatorID", _userId);


            await cmd.ExecuteNonQueryAsync();
        }

        public async Task UpdateInspectionCountsAsync(string _model, string _lotNo, string _userId, int inspectionQty, int okCount)
        {
            const string sql = @"
            UPDATE InspectionData
            SET InspectionQty = @InspectionQty, OkCount = @OkCount
            WHERE PartNo = @PartNo AND LotNo = @LotNo AND OperatorID = @OperatorID";

            using var con = new SqlConnection(_connectionString);
            await con.OpenAsync();

            using var cmd = new SqlCommand(sql, con);
            cmd.Parameters.AddWithValue("@InspectionQty", inspectionQty);
            cmd.Parameters.AddWithValue("@OkCount", okCount);
            cmd.Parameters.AddWithValue("@PartNo", _model);
            cmd.Parameters.AddWithValue("@LotNo", _lotNo);
            cmd.Parameters.AddWithValue("@OperatorID", _userId);

            await cmd.ExecuteNonQueryAsync();
        }

        public List<MeasurementWithConfigModel> GetMasterInspectionWithConfig(string partNo, DateTime? startDate = null, DateTime? endDate = null)
        {
            var list = new List<MeasurementWithConfigModel>();

            // We’ll dynamically unpivot MasterInspection (OL, DE, HD, etc.) to rows using CROSS APPLY
            string query = @"
    SELECT 
        c.Para_No,
        c.Parameter,
        c.Nominal,
        c.RTolPlus,
        c.RTolMinus,
        v.ParameterName AS MeasurementParameter,
        v.MeasurementValue,
        mi.InspectionDate
    FROM MasterInspection mi
    CROSS APPLY (VALUES
        ('OL', mi.OL),
        ('DE', mi.DE),
        ('HD', mi.HD),
        ('GP', mi.GP),
        ('STDG', mi.STDG),
        ('STDU', mi.STDU),
        ('GIR_DIA', mi.GIR_DIA),
        ('STN', mi.STN),
        ('Ovality_SDG', mi.Ovality_SDG),
        ('Ovality_SDU', mi.Ovality_SDU),
        ('Ovality_Head', mi.Ovality_Head),
        ('Stem_Taper', mi.Stem_Taper),
        ('EFRO', mi.EFRO),
        ('Face_Runout', mi.Face_Runout),
        ('SH', mi.SH),
        ('S_RO', mi.S_RO),
        ('DG', mi.DG)
    ) AS v(ParameterName, MeasurementValue)
    INNER JOIN PartConfig c ON c.Parameter = v.ParameterName AND c.Para_No = mi.PartNo
    WHERE mi.PartNo = @PartNo";

            if (startDate.HasValue && endDate.HasValue)
            {
                query += " AND mi.InspectionDate >= @StartDate AND mi.InspectionDate < @EndDate";
            }

            query += " ORDER BY mi.InspectionDate ASC";

            try
            {
                using SqlConnection conn = new(_connectionString);
                using SqlCommand cmd = new(query, conn);

                cmd.Parameters.AddWithValue("@PartNo", partNo);

                if (startDate.HasValue && endDate.HasValue)
                {
                    cmd.Parameters.AddWithValue("@StartDate", startDate.Value.Date);
                    cmd.Parameters.AddWithValue("@EndDate", endDate.Value.Date.AddDays(1));
                }

                conn.Open();
                using SqlDataReader reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    list.Add(new MeasurementWithConfigModel
                    {
                        Para_No = reader["Para_No"].ToString(),
                        Parameter = reader["Parameter"].ToString(),
                        Nominal = Convert.ToDouble(reader["Nominal"]),
                        RTolPlus = Convert.ToDouble(reader["RTolPlus"]),
                        RTolMinus = Convert.ToDouble(reader["RTolMinus"]),
                        MeasurementValue = Convert.ToDouble(reader["MeasurementValue"]),
                        MeasurementDate = Convert.ToDateTime(reader["InspectionDate"])
                    });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("DB Exception: " + ex.Message);
            }

            return list;
        }


        public async Task<List<MeasurementReading>> GetMeasurementReadingsAsync(
       string partNo,
       string lotNo,
       string operatorId,
       DateTime? dateFrom,
       DateTime? dateTo)
        {
            var readings = new List<MeasurementReading>();

            using (SqlConnection con = new SqlConnection(_connectionString))
            {
                await con.OpenAsync();

                // Adjust query to handle null partNo (means all parts)
                string query = @"SELECT *, Status FROM MeasurementReading WHERE (@PartNo IS NULL OR PartNo = @PartNo)";

                if (!string.IsNullOrEmpty(lotNo))
                    query += " AND LotNo = @LotNo";

                if (!string.IsNullOrEmpty(operatorId))
                    query += " AND Operator_ID = @OperatorId";

                if (dateFrom.HasValue)
                    query += " AND MeasurementDate >= @DateFrom";

                if (dateTo.HasValue)
                    query += " AND MeasurementDate <= @DateTo";

                using (SqlCommand cmd = new SqlCommand(query, con))
                {
                    // Pass DBNull.Value if partNo is "All" or null
                    if (string.IsNullOrEmpty(partNo) || partNo.Equals("All", StringComparison.OrdinalIgnoreCase))
                        cmd.Parameters.AddWithValue("@PartNo", DBNull.Value);
                    else
                        cmd.Parameters.AddWithValue("@PartNo", partNo);

                    if (!string.IsNullOrEmpty(lotNo))
                        cmd.Parameters.AddWithValue("@LotNo", lotNo);

                    if (!string.IsNullOrEmpty(operatorId))
                        cmd.Parameters.AddWithValue("@OperatorId", operatorId);

                    if (dateFrom.HasValue)
                        cmd.Parameters.AddWithValue("@DateFrom", dateFrom.Value);

                    if (dateTo.HasValue)
                        cmd.Parameters.AddWithValue("@DateTo", dateTo.Value);

                    using (var reader = await cmd.ExecuteReaderAsync())
                    {
                        while (await reader.ReadAsync())
                        {
                            var reading = new MeasurementReading
                            {
                                Id = reader.GetInt32(reader.GetOrdinal("Id")),
                                PartNo = reader.GetString(reader.GetOrdinal("PartNo")),
                                Operator_ID = reader.GetString(reader.GetOrdinal("Operator_ID")),
                                LotNo = reader.GetString(reader.GetOrdinal("LotNo")),
                                OL = reader.GetDouble(reader.GetOrdinal("OL")),
                                DE = reader.GetDouble(reader.GetOrdinal("DE")),
                                HD = reader.GetDouble(reader.GetOrdinal("HD")),
                                GP = reader.GetDouble(reader.GetOrdinal("GP")),
                                STDG = reader.GetDouble(reader.GetOrdinal("STDG")),
                                STDU = reader.GetDouble(reader.GetOrdinal("STDU")),
                                GIR_DIA = reader.GetDouble(reader.GetOrdinal("GIR_DIA")),
                                STN = reader.GetDouble(reader.GetOrdinal("STN")),
                                Ovality_SDG = reader.GetDouble(reader.GetOrdinal("Ovality_SDG")),
                                Ovality_SDU = reader.GetDouble(reader.GetOrdinal("Ovality_SDU")),
                                Ovality_Head = reader.GetDouble(reader.GetOrdinal("Ovality_Head")),
                                S_RO = reader.GetDouble(reader.GetOrdinal("S_RO")),
                                Stem_Taper = reader.GetDouble(reader.GetOrdinal("Stem_Taper")),
                                EFRO = reader.GetDouble(reader.GetOrdinal("EFRO")),
                                Face_Runout = reader.GetDouble(reader.GetOrdinal("Face_Runout")),
                                SH = reader.GetDouble(reader.GetOrdinal("SH")),
                                DG = reader.GetDouble(reader.GetOrdinal("DG")),
                                MeasurementDate = reader.GetDateTime(reader.GetOrdinal("MeasurementDate")),
                                Status = reader.IsDBNull(reader.GetOrdinal("Status")) ? "Unknown" : reader.GetString(reader.GetOrdinal("Status"))
                            };

                            readings.Add(reading);
                        }
                    }
                }
            }

            return readings;
        }


        public async Task<bool> DeleteMeasurementReadingAsync(int id)
        {
            using (SqlConnection con = new SqlConnection(_connectionString))
            {
                await con.OpenAsync();

                string query = "DELETE FROM MeasurementReading WHERE Id = @Id";

                using (SqlCommand cmd = new SqlCommand(query, con))
                {
                    cmd.Parameters.AddWithValue("@Id", id);

                    int rowsAffected = await cmd.ExecuteNonQueryAsync();

                    // Return true if a record was deleted
                    return rowsAffected > 0;
                }
            }
        }

        public async Task<bool> DeleteLatestMeasurementReadingAsync(string partNo, string lotNo, int operatorId)
        {
            using (SqlConnection con = new SqlConnection(_connectionString))
            {
                await con.OpenAsync();

                string query = @"
            DELETE FROM MeasurementReading
            WHERE Id = (
                SELECT TOP 1 Id FROM MeasurementReading
                WHERE PartNo = @PartNo
                  AND LotNo = @LotNo
                  AND Operator_ID = @OperatorId
                ORDER BY MeasurementDate DESC
            )";

                using (SqlCommand cmd = new SqlCommand(query, con))
                {
                    cmd.Parameters.AddWithValue("@PartNo", partNo);
                    cmd.Parameters.AddWithValue("@LotNo", lotNo);
                    cmd.Parameters.AddWithValue("@OperatorId", operatorId);

                    int rowsAffected = await cmd.ExecuteNonQueryAsync();
                    return rowsAffected > 0;
                }
            }
        }


        public int GetReadingCount()
        {
            const string query = "SELECT ReadingCount FROM ReadingCountTable";
            int readingCount = 0;

            try
            {
                using SqlConnection conn = new(_connectionString);
                using SqlCommand cmd = new(query, conn);

                conn.Open();
                object? result = cmd.ExecuteScalar();

                if (result != null && result != DBNull.Value)
                    readingCount = Convert.ToInt32(result);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error retrieving ReadingCount: {ex.Message}");
            }

            return readingCount;
        }

        // ✅ Update the single ReadingCount value (NO INSERT)
        public void UpdateReadingCount(int newCount)
        {
            const string query = "UPDATE ReadingCountTable SET ReadingCount = @count";

            try
            {
                using SqlConnection conn = new(_connectionString);
                using SqlCommand cmd = new(query, conn);

                cmd.Parameters.AddWithValue("@count", newCount);

                conn.Open();
                cmd.ExecuteNonQuery(); // always updates the existing single row
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error updating ReadingCount: {ex.Message}");
            }
        }



        public void Dispose()
        {
            // Cleanup if needed
        }
    }



    public class PartConfigModel
    {
       // public string SrNo { get; set; }
        public string? Parameter { get; set; }
        public double Nominal { get; set; }
        public double RTolPlus { get; set; }
        public double RTolMinus { get; set; }
        public double YTolPlus { get; set; }
        public double YTolMinus { get; set; }
        public bool IsEnabled { get; set; }
        public int Sign_Change { get; set; }

        public double Compensation { get; set; }
        public string? Para_No { get; set; }

    }

    public class PartReadingDataModel
    {
        public string? Para_No { get; set; }
        public string? Parameter { get; set; }
        public string? ShortName { get; set; }

        public double Nominal { get; set; }
        public double RTolPlus { get; set; }
        public double RTolMinus { get; set; }
        public string? D_Name { get; set; }

        public int Sign_Change { get; set; }

        public double Compensation { get; set; }
    }

    public class ProbeInstallModel
    {
        public string? PartNo { get; set; }
        public string? ProbeId { get; set; }
        public string? Name { get; set; }
        public int Channel { get; set; }

    }

    public class MasterReadingModel
    {
        public string? Para_No { get; set; }
        public string? Parameter { get; set; }
        public double Nominal { get; set; }
        public double RTolPlus { get; set; }
        public double RTolMinus { get; set; }

        public string? D_Name { get; set; }
    }

    public class PartEntryModel
    {
        public string? Para_No { get; set; }
        public string? Para_Name { get; set; }
        public int ActivePart { get; set; }  // 1=active, 0=deactive
    }

    public class MasterReadingProbeReferenceModel
    {
        public string? PartNo { get; set; }
        public string? ProbeId { get; set; }
        public string? ProbeName { get; set; }
        public double Value { get; set; }
        public DateTime LastUpdated { get; set; }
    }

    public class Controls
    {
        public string ? Description { get; set; }

        public int Bit { get; set; }
        public string? Code { get; set; }

    }

    public class InspectionData
    {
        public string PartNo { get; set; }
        public string LotNo { get; set; }
        public string OperatorID { get; set; }
        public int InspectionQty { get; set; }
        public int OkCount { get; set; }
    }

    public class MeasurementWithConfigModel
    {
        public string Para_No { get; set; }
        public string Parameter { get; set; }
        public double Nominal { get; set; }
        public double RTolPlus { get; set; }
        public double RTolMinus { get; set; }
        public double MeasurementValue { get; set; }
        public DateTime MeasurementDate { get; set; }
    }


    public class MeasurementReading
    {
        public int Id { get; set; }
        public string PartNo { get; set; }
        public int TrialNo { get; set; }

        public string Operator_ID { get; set; }
        public string LotNo { get; set; }
        public double OL { get; set; }          // Overall Length
        public double DE { get; set; }          // Datum to End
        public double HD { get; set; }          // Head Diameter
        public double GP { get; set; }          // Groove Position
        public double STDG { get; set; }        // Stem Dia Near Groove
        public double STDU { get; set; }        // Stem Dia Near Undercut
        public double GIR_DIA { get; set; }     // Groove Diameter
        public double STN { get; set; }         // Straightness
        public double Ovality_SDG { get; set; } // Ovality SDG
        public double Ovality_SDU { get; set; } // Ovality SDU
        public double Ovality_Head { get; set; }// Ovality Head
        public double Stem_Taper { get; set; }  // Stem Taper
        public double EFRO { get; set; }        // Face Runout
        public double Face_Runout { get; set; } // Face Runout
        public double SH { get; set; }           // Seat Height
        public double DG { get; set; }           // Seat Height
        public double S_RO { get; set; }           // Seat Height

        public string Status { get; set; }

        public DateTime MeasurementDate { get; set; }
        // ... add other common columns if needed
    }
}
