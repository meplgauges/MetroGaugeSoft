using Microsoft.Data.SqlClient;
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
                    D_Name = reader["D_Name"].ToString(),
                    Nominal = Convert.ToDouble(reader["Nominal"]),
                    RTolPlus = Convert.ToDouble(reader["RTolPlus"]),
                    RTolMinus = Convert.ToDouble(reader["RTolMinus"]),
                    Sign_Change = Convert.ToInt32(reader["Sign_Change"]),
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
                              SELECT PartNo, ProbeName, ParameterName, BoxId, ChannelId
                              FROM ProbeInstallationData
                              WHERE PartNo = @PartNo
                              ORDER BY ProbeName, ChannelId";

            using SqlConnection conn = new(_connectionString);
            using SqlCommand cmd = new(query, conn);
            cmd.Parameters.AddWithValue("@PartNo", partNumber);

            conn.Open();
            using SqlDataReader reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                var item = new ProbeInstallModel
                {
                    PartNo = reader["PartNo"]?.ToString() ?? "",
                    ProbeName = reader["ProbeName"]?.ToString() ?? "",        // ✅ Separate ProbeName column
                    ParameterName = reader["ParameterName"]?.ToString() ?? "", // ✅ Keep ParameterName too
                    BoxId = reader.GetInt32("BoxId"),
                    ChannelId = reader.GetInt32("ChannelId")
                };
                list.Add(item);
            }
            return list;
        }





        public List<PartConfigModel> GetPartConfig(string partNumber)
        {
            var list = new List<PartConfigModel>();

            string query = @"
                    SELECT
                        pc.Parameter,
                        COALESCE(mr.Nominal, pc.Nominal) AS Nominal,
                        pc.RTolPlus,
                        pc.RTolMinus,
                        pc.YTolPlus,
                        pc.YTolMinus,
                        pc.ProbeStatus,
                        pc.Para_No,
                        pc.IsEnabled,
                        pc.Sign_Change,
                        pc.Compensation
                    FROM PartConfig pc
                    LEFT JOIN MasterReadingData mr
                        ON pc.Para_No = mr.Para_No
                       AND pc.Parameter = mr.Parameter
                    WHERE pc.Para_No = @ParaNo
                      AND pc.IsEnabled = 1
                    ORDER BY pc.Parameter";

            using SqlConnection conn = new(_connectionString);
            using SqlCommand cmd = new(query, conn);
            cmd.Parameters.AddWithValue("@ParaNo", partNumber);

            conn.Open();
            using SqlDataReader reader = cmd.ExecuteReader();

            while (reader.Read())
            {
                list.Add(new PartConfigModel
                {
                    Para_No = reader["Para_No"].ToString(),
                    Parameter = reader["Parameter"].ToString(),
                    Nominal = Convert.ToDouble(reader["Nominal"]), // from MasterReadingData if exists
                    RTolPlus = Convert.ToDouble(reader["RTolPlus"]),
                    RTolMinus = Convert.ToDouble(reader["RTolMinus"]),
                    YTolPlus = Convert.ToDouble(reader["YTolPlus"]),
                    YTolMinus = Convert.ToDouble(reader["YTolMinus"]),
                    Sign_Change = Convert.ToInt32(reader["Sign_Change"]),
                    Compensation = Convert.ToDouble(reader["Compensation"])
                    // ProbeStatus → map if needed
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
            p.D_Name,
            COALESCE(m.Nominal, p.Nominal)     AS Nominal,
            COALESCE(m.RTolPlus, p.RTolPlus)   AS RTolPlus,
            COALESCE(m.RTolMinus, p.RTolMinus) AS RTolMinus
        FROM PartConfig p
        LEFT JOIN MasterReadingData m
               ON p.Parameter = m.Parameter
              AND m.Para_No = @PartNumber
        WHERE p.Para_No = @PartNumber
        ORDER BY p.SrNo;
    ";

            using SqlConnection conn = new(_connectionString);
            using SqlCommand cmd = new(query, conn);
            cmd.Parameters.Add("@PartNumber", SqlDbType.VarChar).Value = partNumber;

            conn.Open();
            using SqlDataReader reader = cmd.ExecuteReader();

            while (reader.Read())
            {
                list.Add(new MasterReadingModel
                {
                    Para_No = partNumber,

                    Parameter = reader["Parameter"]?.ToString(),
                    D_Name = reader["D_Name"]?.ToString(),   // ✅ FIXED

                    Nominal = reader["Nominal"] != DBNull.Value ? Convert.ToDouble(reader["Nominal"]) : 0,
                    RTolPlus = reader["RTolPlus"] != DBNull.Value ? Convert.ToDouble(reader["RTolPlus"]) : 0,
                    RTolMinus = reader["RTolMinus"] != DBNull.Value ? Convert.ToDouble(reader["RTolMinus"]) : 0
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



        public List<PartID> GetActiveID()
        {
            var list = new List<PartID>();

            string query = @"
                    SELECT Para_No, ID_Value, BOT_Value
                    FROM Part_Entry
                    WHERE ActivePart = 1";

            using SqlConnection conn = new(_connectionString);
            using SqlCommand cmd = new(query, conn);

            conn.Open();
            using SqlDataReader reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                list.Add(new PartID
                {
                    Para_No = reader["Para_No"].ToString(),
                    ID_Value = Convert.ToInt32(reader["ID_Value"]),
                    BOT_Value = Convert.ToInt32(reader["BOT_Value"])
                });
            }

            return list;
        }

        public PartReadingDataModel? GetOLConfigByPartNumber(string partNumber)
        {
            const string query = @"
 SELECT Parameter, Nominal
 FROM PartConfig
 WHERE Parameter = 'OL'
   AND Para_No = @PartNumber";

            using SqlConnection conn = new(_connectionString);
            using SqlCommand cmd = new(query, conn);
            cmd.Parameters.Add("@PartNumber", SqlDbType.VarChar).Value = partNumber;

            conn.Open();
            using SqlDataReader reader = cmd.ExecuteReader();

            if (reader.Read())
            {
                return new PartReadingDataModel
                {
                    Parameter = reader["Parameter"].ToString(),
                    Nominal = Convert.ToDouble(reader["Nominal"])
                };
            }

            return null;
        }
        public string GetRoboBitByLength(decimal totalLength)
        {
            string roboBit = string.Empty;
            string query = "SELECT RoboBit FROM RoboConfig WHERE TotalLength = @TotalLength";

            using (SqlConnection conn = new SqlConnection(_connectionString))
            using (SqlCommand cmd = new SqlCommand(query, conn))
            {
                cmd.Parameters.AddWithValue("@TotalLength", totalLength);
                conn.Open();
                object result = cmd.ExecuteScalar();
                if (result != null)
                    roboBit = result.ToString();
            }

            return roboBit;
        }


        public List<string> GetAllRoboBits()
        {
            var roboBits = new List<string>();
            string query = "SELECT RoboBit FROM RoboConfig"; // No WHERE clause

            using (SqlConnection conn = new SqlConnection(_connectionString))
            using (SqlCommand cmd = new SqlCommand(query, conn))
            {
                conn.Open();

                using (SqlDataReader reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        if (reader["RoboBit"] != DBNull.Value)
                            roboBits.Add(reader["RoboBit"].ToString());
                    }
                }
            }

            return roboBits;
        }

        public List<PartConfigInfo> GetPartConfigBits(string partNo)
        {
            var list = new List<PartConfigInfo>();
            string query = @"
            SELECT 
                SrNo,
                Para_No,
                ID_Value,
                BOT_Value
            FROM PartConfig
            WHERE Para_No = @Para_No";

            using SqlConnection conn = new(_connectionString);
            using SqlCommand cmd = new(query, conn);
            cmd.Parameters.AddWithValue("@Para_No", partNo);  // ✅ Parameterized by partNo

            conn.Open();
            using SqlDataReader reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                list.Add(new PartConfigInfo
                {
                    SrNo = reader["SrNo"] != DBNull.Value ? Convert.ToInt32(reader["SrNo"]) : 0,
                    Para_No = reader["Para_No"]?.ToString() ?? string.Empty,
                    ID_Value = reader["ID_Value"] != DBNull.Value ? Convert.ToInt32(reader["ID_Value"]) : 0,
                    BOT_Value = reader["BOT_Value"] != DBNull.Value ? Convert.ToInt32(reader["BOT_Value"]) : 0
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

        public List<(string Name, double Min, double Max)> GetMasterProbeRef(string partNo)
        {
            string query = @"
        SELECT Name, MinValue, MaxValue
        FROM MasterReadingProbeReference
        WHERE PartNo = @PartNo";

            var result = new List<(string Name, double Min, double Max)>();

            using SqlConnection conn = new SqlConnection(_connectionString);
            using SqlCommand cmd = new SqlCommand(query, conn);
            cmd.Parameters.AddWithValue("@PartNo", partNo);

            conn.Open();
            using SqlDataReader reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                string name = reader["Name"].ToString() ?? "";

                double min = reader["MinValue"] != DBNull.Value
                    ? Math.Round(Convert.ToDouble(reader["MinValue"]), 3) // round to 3 decimals
                    : 0.0;

                double max = reader["MaxValue"] != DBNull.Value
                    ? Math.Round(Convert.ToDouble(reader["MaxValue"]), 3) // round to 3 decimals
                    : 0.0;

                result.Add((name, min, max));
            }

            return result;
        }



        public void SaveProbeReadings(
    List<ProbeInstallModel> probes,
    string partNo,
    Dictionary<string, (double Min, double Max)> probeValues)   // key = ProbeName
        {
            using SqlConnection conn = new SqlConnection(_connectionString);
            conn.Open();

            foreach (var probe in probes)
            {
                string dictKey = probe.ProbeName;  // "HD001", "SD001"

                if (probeValues.TryGetValue(dictKey, out var range))
                {
                    try
                    {
                        string probeId = probe.ProbeName;  // "HD001"

                        string query = @"
                                        MERGE MasterReadingProbeReference AS target
                                        USING (VALUES (@PartNo, @ProbeId, @Name, @MinValue, @MaxValue)) 
                                               AS source (PartNo, ProbeId, Name, MinValue, MaxValue)
                                        ON (target.PartNo = source.PartNo AND target.ProbeId = source.ProbeId)
                                        WHEN MATCHED THEN 
                                            UPDATE SET 
                                                MinValue    = source.MinValue,
                                                MaxValue    = source.MaxValue,
                                                LastUpdated = GETDATE(),
                                                Name        = source.Name
                                        WHEN NOT MATCHED THEN
                                            INSERT (PartNo, ProbeId, Name, MinValue, MaxValue, LastUpdated)
                                            VALUES (source.PartNo, source.ProbeId, source.Name, 
                                                    source.MinValue, source.MaxValue, GETDATE());
                                        ";

                        using SqlCommand cmd = new(query, conn);
                        cmd.Parameters.AddWithValue("@PartNo", partNo);
                        cmd.Parameters.AddWithValue("@ProbeId", probeId);
                        cmd.Parameters.AddWithValue("@Name", probe.ParameterName);
                        cmd.Parameters.AddWithValue("@MinValue", range.Min);
                        cmd.Parameters.AddWithValue("@MaxValue", range.Max);

                        cmd.ExecuteNonQuery();
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine(
                            $"SQL error saving {dictKey}: {ex.Message}");
                    }
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

            string query = "SELECT * FROM MeasuredData WHERE PartNo = @PartNo";

            if (filterDate.HasValue)
            {
                query += " AND InspectionDate >= @StartDate AND InspectionDate < @EndDate ";
            }

            query += " ORDER BY InspectionDate ASC";

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
        public async Task InsertMasterInspectionAsync(
                          string partNo, string operatorId, string lotNo,
                          decimal stepOd1, decimal stepRunout1,
                          decimal od1, decimal rn1,
                          decimal od2, decimal rn2,
                          decimal od3, decimal rn3,
                          decimal stepOd2, decimal stepRunout2,
                          decimal id1, decimal rn4,
                          decimal id2, decimal rn5,
                          decimal ol,
                          string? status)
        {
            const string query = @"
                            INSERT INTO dbo.MasterInspection
                            (PartNo, Operator_ID, LotNo,
                             [STEP OD1], [STEP RUNOUT-1], [OD-1], [RN-1],
                             [OD-2], [RN-2], [OD-3], [RN-3],
                             [STEP OD2], [STEP RUNOUT-2],
                             [ID-1], [RN-4], [ID-2], [RN-5],
                             OL, Status, InspectionDate)
                            VALUES
                            (@PartNo, @Operator_ID, @LotNo,
                             @STEP_OD1, @STEP_RUNOUT_1, @OD_1, @RN_1,
                             @OD_2, @RN_2, @OD_3, @RN_3,
                             @STEP_OD2, @STEP_RUNOUT_2,
                             @ID_1, @RN_4, @ID_2, @RN_5,
                             @OL, @Status, GETDATE());";

            using var connection = new SqlConnection(_connectionString);
            using var command = new SqlCommand(query, connection);

            command.Parameters.AddWithValue("@PartNo", partNo);
            command.Parameters.AddWithValue("@Operator_ID", operatorId);
            command.Parameters.AddWithValue("@LotNo", lotNo);

            command.Parameters.AddWithValue("@STEP_OD1", stepOd1);
            command.Parameters.AddWithValue("@STEP_RUNOUT_1", stepRunout1);
            command.Parameters.AddWithValue("@OD_1", od1);
            command.Parameters.AddWithValue("@RN_1", rn1);
            command.Parameters.AddWithValue("@OD_2", od2);
            command.Parameters.AddWithValue("@RN_2", rn2);
            command.Parameters.AddWithValue("@OD_3", od3);
            command.Parameters.AddWithValue("@RN_3", rn3);
            command.Parameters.AddWithValue("@STEP_OD2", stepOd2);
            command.Parameters.AddWithValue("@STEP_RUNOUT_2", stepRunout2);
            command.Parameters.AddWithValue("@ID_1", id1);
            command.Parameters.AddWithValue("@RN_4", rn4);
            command.Parameters.AddWithValue("@ID_2", id2);
            command.Parameters.AddWithValue("@RN_5", rn5);
            command.Parameters.AddWithValue("@OL", ol);

            command.Parameters.AddWithValue("@Status", (object?)status ?? DBNull.Value);

            await connection.OpenAsync();
            var rows = await command.ExecuteNonQueryAsync();
            if (rows == 0)
                throw new Exception("Insert failed: No rows were affected.");
        }


        public async Task InsertMeasurementReadingAsync(
    string partNo, string operatorId, string lotNo,
                          decimal stepOd1, decimal stepRunout1,
                          decimal od1, decimal rn1,
                          decimal od2, decimal rn2,
                          decimal od3, decimal rn3,
                          decimal stepOd2, decimal stepRunout2,
                          decimal id1, decimal rn4,
                          decimal id2, decimal rn5,
                          decimal ol,
                          string? status)
        {
            const string query = @"
                            INSERT INTO dbo.MeasuredData
                            (PartNo, Operator_ID, LotNo,
                             [STEP OD1], [STEP RUNOUT-1], [OD-1], [RN-1],
                             [OD-2], [RN-2], [OD-3], [RN-3],
                             [STEP OD2], [STEP RUNOUT-2],
                             [ID-1], [RN-4], [ID-2], [RN-5],
                             OL, Status, InspectionDate)
                            VALUES
                            (@PartNo, @Operator_ID, @LotNo,
                             @STEP_OD1, @STEP_RUNOUT_1, @OD_1, @RN_1,
                             @OD_2, @RN_2, @OD_3, @RN_3,
                             @STEP_OD2, @STEP_RUNOUT_2,
                             @ID_1, @RN_4, @ID_2, @RN_5,
                             @OL, @Status, GETDATE());";

            using var connection = new SqlConnection(_connectionString);
            using var command = new SqlCommand(query, connection);

            command.Parameters.AddWithValue("@PartNo", partNo);
            command.Parameters.AddWithValue("@Operator_ID", operatorId);
            command.Parameters.AddWithValue("@LotNo", lotNo);

            command.Parameters.AddWithValue("@STEP_OD1", stepOd1);
            command.Parameters.AddWithValue("@STEP_RUNOUT_1", stepRunout1);
            command.Parameters.AddWithValue("@OD_1", od1);
            command.Parameters.AddWithValue("@RN_1", rn1);
            command.Parameters.AddWithValue("@OD_2", od2);
            command.Parameters.AddWithValue("@RN_2", rn2);
            command.Parameters.AddWithValue("@OD_3", od3);
            command.Parameters.AddWithValue("@RN_3", rn3);
            command.Parameters.AddWithValue("@STEP_OD2", stepOd2);
            command.Parameters.AddWithValue("@STEP_RUNOUT_2", stepRunout2);
            command.Parameters.AddWithValue("@ID_1", id1);
            command.Parameters.AddWithValue("@RN_4", rn4);
            command.Parameters.AddWithValue("@ID_2", id2);
            command.Parameters.AddWithValue("@RN_5", rn5);
            command.Parameters.AddWithValue("@OL", ol);

            command.Parameters.AddWithValue("@Status", (object?)status ?? DBNull.Value);

            await connection.OpenAsync();
            var rows = await command.ExecuteNonQueryAsync();
            if (rows == 0)
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


        public async Task<InspectionData?> SelectInspectionDataAsync(string? _model, string? _lotNo, string? _userId)
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
        public async Task<List<string>> GetLotNumbersByPartAndDateRangeAsync(
     string? partNo,
     DateTime? dateFrom,
     DateTime? dateTo)
        {
            var lotNumbers = new List<string>();

            using (var con = new SqlConnection(_connectionString))
            {
                await con.OpenAsync();

                string query = @"
                                SELECT DISTINCT LotNo
                                FROM dbo.MeasuredData
                                WHERE 1 = 1";

                if (!string.IsNullOrEmpty(partNo))
                    query += " AND PartNo = @PartNo";

                if (dateFrom.HasValue)
                    query += " AND InspectionDate >= @DateFrom";

                if (dateTo.HasValue)
                    query += " AND InspectionDate <= @DateTo";

                using (var cmd = new SqlCommand(query, con))
                {
                    if (!string.IsNullOrEmpty(partNo))
                        cmd.Parameters.AddWithValue("@PartNo", partNo);

                    if (dateFrom.HasValue)
                        cmd.Parameters.AddWithValue("@DateFrom", dateFrom.Value);

                    if (dateTo.HasValue)
                        cmd.Parameters.AddWithValue("@DateTo", dateTo.Value);

                    using (var reader = await cmd.ExecuteReaderAsync())
                    {
                        int lotNoOrdinal = reader.GetOrdinal("LotNo");

                        while (await reader.ReadAsync())
                        {
                            if (!reader.IsDBNull(lotNoOrdinal))
                                lotNumbers.Add(reader.GetString(lotNoOrdinal));
                        }
                    }
                }
            }

            return lotNumbers;
        }


        public async Task<List<string>> GetOperatorsByPartAndDateRangeAsync(string? partNo, DateTime? dateFrom, DateTime? dateTo)
        {
            var operators = new List<string>();

            using (SqlConnection con = new SqlConnection(_connectionString))
            {
                await con.OpenAsync();

                string query = @"SELECT DISTINCT Operator_ID FROM MeasuredData WHERE 1=1";

                if (!string.IsNullOrEmpty(partNo))
                    query += " AND PartNo = @PartNo";

                if (dateFrom.HasValue)
                    query += " AND " +
                        "InspectionDate >= @DateFrom";

                if (dateTo.HasValue)
                    query += " AND InspectionDate <= @DateTo";

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

        public async Task InsertInspectionDataAsync(string? _model, string? _lotNo, string? _userId)
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

        public async Task UpdateInspectionCountsAsync(string? _model, string? _lotNo, string? _userId, int inspectionQty, int okCount)
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

        public List<MeasurementWithConfigModel> GetMasterInspectionWithConfig(string? partNo, DateTime? startDate = null, DateTime? endDate = null)
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
                     string? partNo,
                     string? lotNo,
                     string? operatorId,
                     DateTime? dateFrom,
                     DateTime? dateTo)
        {
            var readings = new List<MeasurementReading>();

            using (var con = new SqlConnection(_connectionString))
            {
                await con.OpenAsync();

                string query = @"
                            SELECT *
                            FROM dbo.MeasuredData
                            WHERE (@PartNo IS NULL OR PartNo = @PartNo)";

                if (!string.IsNullOrEmpty(lotNo))
                    query += " AND LotNo = @LotNo";

                if (!string.IsNullOrEmpty(operatorId))
                    query += " AND Operator_ID = @OperatorId";

                if (dateFrom.HasValue)
                    query += " AND InspectionDate >= @DateFrom";

                if (dateTo.HasValue)
                    query += " AND InspectionDate <= @DateTo";

                using (var cmd = new SqlCommand(query, con))
                {
                    if (string.IsNullOrEmpty(partNo) ||
                        partNo.Equals("All", StringComparison.OrdinalIgnoreCase))
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
                            var r = new MeasurementReading
                            {
                                Id = reader.GetInt32(reader.GetOrdinal("ID")),
                                PartNo = reader.GetString(reader.GetOrdinal("PartNo")),
                                Operator_ID = reader.GetString(reader.GetOrdinal("Operator_ID")),
                                LotNo = reader.GetString(reader.GetOrdinal("LotNo")),

                                StepOd1 = reader.GetDecimal(reader.GetOrdinal("STEP OD1")),
                                StepRunout1 = reader.GetDecimal(reader.GetOrdinal("STEP RUNOUT-1")),
                                Od1 = reader.GetDecimal(reader.GetOrdinal("OD-1")),
                                Rn1 = reader.GetDecimal(reader.GetOrdinal("RN-1")),
                                Od2 = reader.GetDecimal(reader.GetOrdinal("OD-2")),
                                Rn2 = reader.GetDecimal(reader.GetOrdinal("RN-2")),
                                Od3 = reader.GetDecimal(reader.GetOrdinal("OD-3")),
                                Rn3 = reader.GetDecimal(reader.GetOrdinal("RN-3")),
                                StepOd2 = reader.GetDecimal(reader.GetOrdinal("STEP OD2")),
                                StepRunout2 = reader.GetDecimal(reader.GetOrdinal("STEP RUNOUT-2")),
                                Id1 = reader.GetDecimal(reader.GetOrdinal("ID-1")),
                                Rn4 = reader.GetDecimal(reader.GetOrdinal("RN-4")),
                                Id2 = reader.GetDecimal(reader.GetOrdinal("ID-2")),
                                Rn5 = reader.GetDecimal(reader.GetOrdinal("RN-5")),
                                Ol = reader.GetDecimal(reader.GetOrdinal("OL")),

                                MeasurementDate = reader.GetDateTime(reader.GetOrdinal("InspectionDate")),
                                Status = reader.IsDBNull(reader.GetOrdinal("Status"))
                                    ? "Unknown"
                                    : reader.GetString(reader.GetOrdinal("Status"))
                            };

                            readings.Add(r);
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

                string query = "DELETE FROM MeasuredData WHERE ID = @ID";

                using (SqlCommand cmd = new SqlCommand(query, con))
                {
                    cmd.Parameters.AddWithValue("@ID", id);

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
                                DELETE FROM MeasuredData
                                WHERE Id = (
                                    SELECT TOP 1 ID FROM MeasuredData
                                    WHERE PartNo = @PartNo
                                      AND LotNo = @LotNo
                                      AND Operator_ID = @OperatorId
                                    ORDER BY InspectionDate DESC
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
        public string? PartNo { get; set; } = "";
        public string? ParameterName { get; set; } = "";  // ✅ "Head Diameter"
        public int BoxId { get; set; }                   // ✅ 1, 2, 3...
        public int ChannelId { get; set; }               // ✅ 1, 2, 3, 4

        // Keep old properties for backward compatibility (if needed)
        public string? ProbeId { get; set; } = "";
        public string ProbeName { get; set; } = "";           // ✅ Unique ID column
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
        public string? Description { get; set; }

        public int Bit { get; set; }
        public string? Code { get; set; }

    }

    public class InspectionData
    {
        public string? PartNo { get; set; }
        public string? LotNo { get; set; }
        public string? OperatorID { get; set; }
        public int InspectionQty { get; set; }
        public int OkCount { get; set; }
    }

    public class MeasurementWithConfigModel
    {
        public string? Para_No { get; set; }
        public string? Parameter { get; set; }
        public double Nominal { get; set; }
        public double RTolPlus { get; set; }
        public double RTolMinus { get; set; }
        public double MeasurementValue { get; set; }
        public DateTime MeasurementDate { get; set; }
    }


    public class PartConfigInfo
    {
        public int SrNo { get; set; }
        public string Para_No { get; set; } = string.Empty;
        public int ID_Value { get; set; }    // 0=No Select, 1=ID1, 2=ID2, 3=ID3
        public int BOT_Value { get; set; }   // 0=No Select, 1=CYC1, 2=CYC2, 3=CYC3
    }


    public class PartID
    {
        public string Para_No { get; set; }
        public int ID_Value { get; set; }
        public int BOT_Value { get; set; }
    }


    public class MeasurementReading
    {
        public int Id { get; set; }
        public string PartNo { get; set; } = "";
        public string Operator_ID { get; set; } = "";
        public string LotNo { get; set; } = "";

        public decimal StepOd1 { get; set; }
        public decimal StepRunout1 { get; set; }
        public decimal Od1 { get; set; }
        public decimal Rn1 { get; set; }
        public decimal Od2 { get; set; }
        public decimal Rn2 { get; set; }
        public decimal Od3 { get; set; }
        public decimal Rn3 { get; set; }
        public decimal StepOd2 { get; set; }
        public decimal StepRunout2 { get; set; }
        public decimal Id1 { get; set; }
        public decimal Rn4 { get; set; }
        public decimal Id2 { get; set; }
        public decimal Rn5 { get; set; }
        public decimal Ol { get; set; }

        public int TrialNo { get; set; }
        public DateTime MeasurementDate { get; set; }
        public string Status { get; set; } = "";
    }

}
