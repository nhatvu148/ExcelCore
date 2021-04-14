﻿using OfficeOpenXml;
using System;
using System.IO;
using System.Reflection;
using MySql.Data.MySqlClient;
using Newtonsoft.Json;
using System.Collections.Generic;

namespace ExcelCore
{
    class Program
    {
        static void Main(string[] args)
        {
            string host = "localhost";
            string user = "root";
            string password = "123456789";
            string database = "jmu";
            string shipInfoID = "1";
            string startTime = "";
            string endTime = "";
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
            }

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            string pathDirectory = $"{Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)}";

            var fileSource = new FileInfo($"{pathDirectory}/sample.xlsx");
            var fileGyroSource = new FileInfo($"{pathDirectory}/Gyro.xlsx");
            var fileDestination = new FileInfo($"{pathDirectory}/../output/応力・加速度グラフ.xlsx");
            var fileGyroDestination = new FileInfo($"{pathDirectory}/../output/Gyroグラフ.xlsx");

            string cs = @$"server={host};userid={user};password={password};database={database}";

            using var con = new MySqlConnection(cs);
            con.Open();

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
        }
    }
}
