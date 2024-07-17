using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data.SqlClient;
using System.Dynamic;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices.ComTypes;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using Xbim.Common;
using Xbim.Ifc2x3.Kernel;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel;

namespace HelloWorld
{
    public class SQLDBConnect
    {
        private SqlConnection connect;

        private string connectionString;

        static public int NumOfAverageMaterials = 0;

        public bool ClearAverageData()
        {
            using (SqlTransaction transaction = connect.BeginTransaction())
            {
                string[] tableNames = { "AverageDataTable", "AverageLocationTable", "AverageManufacturerTable", "AverageMaterialTable" };

                try
                {
                    foreach (string tableName in tableNames)
                    {
                        string deleteCommandText = "DELETE FROM [dbo].[{tableName}]";
                        // Create SqlCommand
                        using (SqlCommand command = new SqlCommand(deleteCommandText, connect, transaction))
                        {
                            // Execute the command
                            command.ExecuteNonQuery();
                        }
                    }
                    
                    // Commit the transaction
                    transaction.Commit();

                    MessageBox.Show("Total " + NumOfAverageMaterials.ToString() + " rows have been deleted");

                    NumOfAverageMaterials = 0;

                    return true;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("An error occurred: " + ex.Message);
                    // Rollback the transaction if an error occurs
                    transaction.Rollback();

                    return false;
                }
            }
        }

        public bool ConnectDB(string connectionString)
        {
            try
            {
                connect = new SqlConnection(connectionString);

                if (connect.State == System.Data.ConnectionState.Closed)
                {
                    connect.Open();

                    NumOfAverageMaterials = GetAvergaeMaterialNum() + 1;

                    MessageBox.Show("Connect to database successfully!");
                }
                else
                {
                    connect.Close();

                    connect.Open();

                    NumOfAverageMaterials = GetAvergaeMaterialNum() + 1;

                    MessageBox.Show("Database is reconnected!");
                }

                return true;
            }
            catch
            {
                MessageBox.Show("Connect to database failed!");

                return false;
            }
        }

        private int GetAvergaeMaterialNum()
        {
            string queryString = "SELECt COUNT(*) FROM [dbo].[AverageMaterialTable]";

            SqlCommand command = new SqlCommand(queryString, connect);

            int count = (int)command.ExecuteScalar();

            return count;
        }

        public void CloseDB()
        {
            if (connect.State == System.Data.ConnectionState.Open)
            {
                connect.Close();
            }
        }

        public SqlCommand Query(string query)
        {
            SqlCommand command = new SqlCommand(query, connect);

            return command;
        }

        internal (bool,Guid) InsertData(List<(Guid, int, string)> records, string materialType, string year)
        {
            Guid averageMaterialID = Guid.NewGuid();

            NumOfAverageMaterials++;

            DateTime timeStamp = DateTime.Now;

            string averageDataRecord = "No." + NumOfAverageMaterials.ToString() + "_" + materialType + "_" + timeStamp;

            var transaction = connect.BeginTransaction();

            List<string> cities = new List<string>();

            List<int> manufacturers = new List<int>();

            foreach(var record in records)
            {
                cities.Add(record.Item3);

                manufacturers.Add(record.Item2);
            }

            cities = cities.Distinct().ToList();

            manufacturers = manufacturers.Distinct().ToList();

            try
            {
                SqlCommand enableIdentityInsertCommand = new SqlCommand("SET IDENTITY_INSERT [dbo].[AverageMaterialTable] ON", connect, transaction);

                enableIdentityInsertCommand.ExecuteNonQuery();

                using (var insertCommand = new SqlCommand("INSERT INTO [dbo].[AverageMaterialTable] ([No.],[Material_ID],[Material_Name],[Material_Type],[NumOfRecords],[NumOfRegions],[NumOfManufacturers],[StartDate],[EndDate],[GenerateTime]) VALUES (@No,@Material_ID,@Material_Name, @Material_Type,@NumOfRecords,@NumOfRegions,@NumofManufacturers,@StartDate,@EndDate,@GenerateTime)", connect, transaction))
                {
                    // Set parameter values
                    insertCommand.Parameters.AddWithValue("@No", NumOfAverageMaterials);

                    insertCommand.Parameters.AddWithValue("@Material_ID", averageMaterialID);

                    insertCommand.Parameters.AddWithValue("@Material_Name", averageDataRecord);

                    insertCommand.Parameters.AddWithValue("@Material_Type", materialType);

                    insertCommand.Parameters.AddWithValue("@NumOfRecords", records.Count);

                    insertCommand.Parameters.AddWithValue("@NumOfRegions", cities.Count);//0 represents no limits

                    insertCommand.Parameters.AddWithValue("@NumofManufacturers", manufacturers.Count);//0 represents 0

                    string startDate = null;

                    string endDate = null;
                    
                    if(year=="XXX")
                    {
                        startDate=timeStamp.Date.ToString("yyyy-01-01");

                        endDate = timeStamp.Date.ToString("yyyy-12-31");
                    }else
                    {
                        startDate = $"{year}-01-01";

                        endDate = $"{year}-12-31"; ;
                    }

                    insertCommand.Parameters.AddWithValue("@StartDate", startDate);

                    insertCommand.Parameters.AddWithValue("@EndDate", endDate);

                    insertCommand.Parameters.AddWithValue("@GenerateTime", timeStamp);

                    // Execute INSERT command
                    insertCommand.ExecuteNonQuery();
                }

                foreach (var manufacturer in manufacturers)
                {
                    using (var insertCommand = new SqlCommand("INSERT INTO [dbo].[AverageManufacturerTable] ([Manufacturer_ID], [AverageMaterial_ID]) VALUES (@Manufacturer_ID, @AverageMaterial_ID)", connect, transaction))
                    {
                        // Set parameter values
                        insertCommand.Parameters.AddWithValue("@AverageMaterial_ID", averageMaterialID);

                        insertCommand.Parameters.AddWithValue("@Manufacturer_ID", manufacturer);

                        // Execute INSERT command
                        insertCommand.ExecuteNonQuery();
                    }
                }

                foreach (var city in cities)
                {
                    using (var insertCommand = new SqlCommand("INSERT INTO [dbo].[AverageLocationTable] ([City_Name], [AverageMaterial_ID]) VALUES (@City_Name, @AverageMaterial_ID)", connect, transaction))
                    {
                        // Set parameter values
                        insertCommand.Parameters.AddWithValue("@AverageMaterial_ID", averageMaterialID);

                        insertCommand.Parameters.AddWithValue("@City_Name", city);

                        // Execute INSERT command
                        insertCommand.ExecuteNonQuery();
                    }
                }

                foreach (var record in records)
                {
                    using (var insertCommand = new SqlCommand("INSERT INTO [dbo].[AverageDataTable] ([Material_ID], [AverageMaterial_ID]) VALUES (@Material_ID, @AverageMaterial_ID)", connect, transaction))
                    {
                        // Set parameter values

                        insertCommand.Parameters.AddWithValue("@Material_ID", record.Item1);

                        insertCommand.Parameters.AddWithValue("@AverageMaterial_ID", averageMaterialID);

                        // Execute INSERT command
                        insertCommand.ExecuteNonQuery();
                    }
                }
                transaction.Commit();

                return (true, averageMaterialID);
            }
            catch (Exception ex)
            {
                transaction.Rollback();

                return (false, new Guid());
            }
        }

        internal void CleanUp(DateTime timeStamp)
        {

            string sql ="Select Material_ID from [dbo].[AverageMaterialTable] where [GenerateTime] > @timeStamp";

            SqlCommand command = new SqlCommand(sql, connect);

            command.Parameters.AddWithValue("@timeStamp", timeStamp);

            SqlDataReader reader = command.ExecuteReader();

            List<Guid> materialIDs = new List<Guid>();

            while (reader.Read())
            {
                materialIDs.Add(reader.GetGuid(0));
            }

            reader.Close();

            foreach (var materialID in materialIDs)
            {
                DeleteData(materialID);
                DeleteManufacturer(materialID);
                DeleteLocations(materialID);
                DeleteMaterial(materialID);
            }

        }

        private void DeleteMaterial(Guid materialID)
        {
            string sql = $"DELETE FROM [dbo].[AverageMaterialTable] WHERE [Material_ID] ='{materialID.ToString()}'";

            var transaction = connect.BeginTransaction();

            SqlCommand command = new SqlCommand(sql, connect, transaction);

            command.ExecuteNonQuery();

            transaction.Commit();

        }

        private void DeleteLocations(Guid materialID)
        {
            string sql = $"DELETE FROM [dbo].[AverageLocationTable] WHERE [AverageMaterial_ID] ='{materialID.ToString()}'";

            var transaction = connect.BeginTransaction();

            SqlCommand command = new SqlCommand(sql, connect, transaction);

            command.ExecuteNonQuery();

            transaction.Commit();
        }

        private void DeleteManufacturer(Guid materialID)
        {
            string sql = $"DELETE FROM [dbo].[AverageManufacturerTable] WHERE [AverageMaterial_ID] ='{materialID.ToString()}'";

            var transaction = connect.BeginTransaction();

            SqlCommand command = new SqlCommand(sql, connect, transaction);

            command.ExecuteNonQuery();

            transaction.Commit();
        }

        private void DeleteData(Guid materialID)
        {
            string sql = $"DELETE FROM [dbo].[AverageDataTable] WHERE [AverageMaterial_ID] ='{materialID.ToString()}'";

            var transaction = connect.BeginTransaction();

            SqlCommand command = new SqlCommand(sql, connect, transaction);

            command.ExecuteNonQuery();

            transaction.Commit();
        }
    }
}
