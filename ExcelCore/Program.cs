using OfficeOpenXml;
using System;
using System.IO;
using System.Reflection;
using MySql.Data.MySqlClient;
using Newtonsoft.Json;
using System.Collections.Generic;
using System.Threading.Tasks;
// using System.Diagnostics;

namespace ExcelCore
{
    class Program
    {
        static void Main(string[] args)
        {
            // Stopwatch stopwatch = new Stopwatch();
            // stopwatch.Start();

            string pathDirectory = $"{Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)}";

            string host = "localhost";
            string user = "root";
            string password = "123456789";
            string database = "jmu";
            string shipInfoID = "1";
            string startTime = "";
            string endTime = "";
            string outDir = $"{pathDirectory}/../output";
            if (args == null || args.Length == 0)
            {
                // no arguments
                //startTime = "2021-01-15 23:40:00";
                //endTime = "2021-01-25 23:00:00";
            }
            else
            {
                host = Convert.ToString(args[0]);
                user = Convert.ToString(args[1]);
                password = Convert.ToString(args[2]);
                database = Convert.ToString(args[3]);
                shipInfoID = Convert.ToString(args[4]);
                startTime = Convert.ToString(args[5]);
                endTime = Convert.ToString(args[6]);
                outDir = Convert.ToString(args[7]);
            }

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            var fileSource = new FileInfo($"{pathDirectory}/sample.xlsx");
            var fileGyroSource = new FileInfo($"{pathDirectory}/Gyro.xlsx");
            var fileWaveSource = new FileInfo($"{pathDirectory}/Wave.xlsx");
            var fileDestination = new FileInfo($"{outDir}/応力・加速度グラフ.xlsx");
            var fileGyroDestination = new FileInfo($"{outDir}/Gyroグラフ.xlsx");
            var fileWaveDestination = new FileInfo($"{outDir}/Waveグラフ.xlsx");

            string cs = @$"server={host};userid={user};password={password};database={database}";

            Parallel.Invoke(() =>
            {
                using var con = new MySqlConnection(cs);
                con.Open();

                // Get SQL data for Main charts
                string sql = startTime != "" && endTime != "" ?
                    $"SELECT NumofMeasurePoint, MeasurePointData FROM statistics.state_statistics WHERE ShipInfo_ID='{shipInfoID}' AND datetime BETWEEN '{startTime}' AND '{endTime}'"
                    :
                    $"SELECT NumofMeasurePoint, MeasurePointData FROM statistics.state_statistics WHERE ShipInfo_ID='{shipInfoID}'"
                    ;

                using var cmd = new MySqlCommand(sql, con);

                using MySqlDataReader rdr = cmd.ExecuteReader();

                int numofMeasurePoint = 0;
                List<double[]> menrList = new List<double[]>();
                List<double[]> devlList = new List<double[]>();

                while (rdr.Read())
                {
                    numofMeasurePoint = JsonConvert.DeserializeObject<int>(rdr.GetString(0));

                    double[] m1 = JsonConvert.DeserializeObject<double[]>(rdr.GetString(1));
                    menrList.Add(m1[0..(numofMeasurePoint + 2)]);
                    devlList.Add(m1[(numofMeasurePoint + 2)..(2 * numofMeasurePoint + 4)]);
                }
                rdr.Close();

                // Create Main charts
                using (var excelFileSource = new ExcelPackage(fileSource))
                using (var excelFileDestination = new ExcelPackage(fileDestination))
                {
                    var menrWorksheetSource = excelFileSource.Workbook.Worksheets[0];
                    for (int i = 1; i <= menrList.Count; i++)
                    {
                        for (int j = 1; j <= numofMeasurePoint + 2; j++)
                        {
                            menrWorksheetSource.Cells[i + 1, j].Value = menrList[i - 1][j - 1];
                        }
                    }

                    var devlWorksheetSource = excelFileSource.Workbook.Worksheets[1];
                    for (int i = 1; i <= devlList.Count; i++)
                    {
                        for (int j = 1; j <= numofMeasurePoint + 2; j++)
                        {
                            devlWorksheetSource.Cells[i + 1, j].Value = devlList[i - 1][j - 1];
                        }
                    }

                    excelFileSource.SaveAs(fileDestination);
                }
            }, () =>
            {
                using var con = new MySqlConnection(cs);
                con.Open();

                // Get SQL data for Gyro charts
                string sqlGyro = startTime != "" && endTime != "" ?
                    $"SELECT datetime, Roll_Max, Pitch_Max, Yaw_Max FROM statistics.gyro WHERE ShipInfo_ID='{shipInfoID}' AND datetime BETWEEN '{startTime}' AND '{endTime}'"
                    :
                    $"SELECT datetime, Roll_Max, Pitch_Max, Yaw_Max FROM statistics.gyro WHERE ShipInfo_ID='{shipInfoID}'"
                    ;
                using var cmdGyro = new MySqlCommand(sqlGyro, con);
                using MySqlDataReader rdrGyro = cmdGyro.ExecuteReader();

                List<string> dateList = new List<string>();
                List<double> rollList = new List<double>();
                List<double> pitchList = new List<double>();
                List<double> yawList = new List<double>();
                while (rdrGyro.Read())
                {
                    dateList.Add(rdrGyro.GetString(0));
                    rollList.Add(JsonConvert.DeserializeObject<double>(rdrGyro.GetString(1)));
                    pitchList.Add(JsonConvert.DeserializeObject<double>(rdrGyro.GetString(2)));
                    yawList.Add(JsonConvert.DeserializeObject<double>(rdrGyro.GetString(3)));
                }
                rdrGyro.Close();

                // Create Gyro charts
                using (var excelFileSource = new ExcelPackage(fileGyroSource))
                {
                    var gyroWorksheetSource = excelFileSource.Workbook.Worksheets[0];
                    for (int i = 0; i < dateList.Count; i++)
                    {
                        gyroWorksheetSource.Cells[i + 2, 1].Value = dateList[i];
                        gyroWorksheetSource.Cells[i + 2, 2].Value = rollList[i];
                        gyroWorksheetSource.Cells[i + 2, 3].Value = pitchList[i];
                        gyroWorksheetSource.Cells[i + 2, 4].Value = yawList[i];
                    }

                    excelFileSource.SaveAs(fileGyroDestination);
                }
            }, () =>
            {
                using var con = new MySqlConnection(cs);
                con.Open();

                // Get SQL data for Wave charts
                string sqlWave = startTime != "" && endTime != "" ?
                    $"SELECT datetime, WaveHeight, WavePeriod FROM statistics.waves WHERE ShipInfo_ID='{shipInfoID}' AND datetime BETWEEN '{startTime}' AND '{endTime}'"
                    :
                    $"SELECT datetime, WaveHeight, WavePeriod FROM statistics.waves WHERE ShipInfo_ID='{shipInfoID}'"
                    ;
                using var cmdWave = new MySqlCommand(sqlWave, con);
                using MySqlDataReader rdrWave = cmdWave.ExecuteReader();

                List<string> dateList2 = new List<string>();
                List<double> waveHList = new List<double>();
                List<double> wavePList = new List<double>();
                while (rdrWave.Read())
                {
                    dateList2.Add(rdrWave.GetString(0));
                    waveHList.Add(JsonConvert.DeserializeObject<double>(rdrWave.GetString(1)));
                    wavePList.Add(JsonConvert.DeserializeObject<double>(rdrWave.GetString(2)));
                }
                rdrWave.Close();

                // Create Wave charts
                using (var excelFileSource = new ExcelPackage(fileWaveSource))
                {
                    var waveWorksheetSource = excelFileSource.Workbook.Worksheets[0];
                    for (int i = 0; i < dateList2.Count; i++)
                    {
                        waveWorksheetSource.Cells[i + 2, 1].Value = dateList2[i];
                        waveWorksheetSource.Cells[i + 2, 2].Value = waveHList[i];
                        waveWorksheetSource.Cells[i + 2, 3].Value = wavePList[i];
                    }

                    excelFileSource.SaveAs(fileWaveDestination);
                }
            });

            // stopwatch.Stop();
            // Console.WriteLine("Time elapsed: {0}", stopwatch.Elapsed);
        }
    }
}
