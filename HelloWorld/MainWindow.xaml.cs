using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Navigation;
using Xbim.Presentation.XplorerPluginSystem;
using System.IO;
using CsvHelper;
using System.Globalization;
using Newtonsoft.Json;
using System.Diagnostics;
using System.Threading;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Runtime.InteropServices;
using XbimXplorer;
using Xbim.Ifc;
using Xbim.Ifc4.Interfaces;
using System.Collections;
using Xbim.Ifc4.ProductExtension;
using Xbim.Ifc4.SharedBldgElements;
using Xbim.Ifc4.StructuralElementsDomain;
using Xbim.Ifc4.PlumbingFireProtectionDomain;
using Xbim.Ifc4.ElectricalDomain;
using Xbim.Ifc4.Kernel;
using Xbim.Ifc4.PropertyResource;
using Xbim.Common;
using Xbim.Presentation;
using System.ComponentModel;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel;
using System.Data.SqlClient;
using HelixToolkit.Wpf;
using System.Net;
using System.Runtime.InteropServices.ComTypes;
using Xbim.Ifc4.SharedBldgServiceElements;
using Xbim.Ifc4.BuildingControlsDomain;
using Xbim.Ifc4.HvacDomain;
using System.Windows.Media.Imaging;
using Xbim.Ifc4.QuantityResource;
using Xbim.Ifc4.MeasureResource;
using Xbim.IO.Xml.BsConf;
using System.Windows.Media.Media3D;
using Xbim.Ifc4.SharedFacilitiesElements;
using System.Runtime.Remoting.Contexts;
using Xbim.Ifc4.MaterialResource;
using System.Collections.Concurrent;
using System.Text.RegularExpressions;
using System.Windows.Media.Effects;
using System.Xml.Linq;
using Xbim.Ifc2x3.MeasureResource;

namespace HelloWorld
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>

    [XplorerUiElement(PluginWindowUiContainerEnum.LayoutAnchorable, PluginWindowActivation.OnMenu, "LCA Calculator")]

    public partial class MainWindow : IXbimXplorerPluginWindow
    {

        #region Members

        List<Record> A1_A3_Min;

        List<Record> A1_A3_Max;

        List<Record> A4_Min;

        List<Record> A4_Max;

        List<Record> A5_Min;

        List<Record> A5_Max;

        List<Record> A1_A3_Avg;

        List<Record> A4_Avg;

        List<Record> A5_Avg;

        List<Record> A1_A3_SD;

        List<Record> A4_SD;

        List<Record> A5_SD;

        List<Record> All_Stage_Min;

        List<Record> All_Stage_Max;

        List<Record> All_Stage_Avg;

        List<Record> All_Stage_SD;

        List<Record> Total_Min;

        List<Record> Total_Max;

        List<Record> Total_Avg;

        List<Record> Total_SD;

        string status_1;

        string status_2;

        IXbimXplorerPluginMasterWindow _xpWindow;

        public string WindowTitle => "LCA Calculator";

        //run a new process to run the server, close it when the window is closed
        static Process serverProcess;

        bool serverStarted = false;

        DistanceCalculator distanceCalculator = new DistanceCalculator("5b3ce3597851110001cf6248ff1788b40b624a5c8ef99b7e4e089734");

        Thread thread;       

        string openLCAPath = null;

        SQLDBConnect connect = new SQLDBConnect();

        List<(string,string)> failedMaterials=new List<(string,string)>();

        List<IfcElement> SelectedElements;

        List<TreeNode<string>> elementList = new List<TreeNode<string>>();
        // Selection
        public EntitySelection Selection
        {
            get { return (EntitySelection)GetValue(SelectionProperty); }
            set { SetValue(SelectionProperty, value); }
        }

        public static DependencyProperty SelectionProperty =
            DependencyProperty.Register("Selection", typeof(EntitySelection), typeof(MainWindow), new FrameworkPropertyMetadata(null, FrameworkPropertyMetadataOptions.Inherits, OnPropertyChanged));

        // SelectedEntity
        public IPersistEntity SelectedEntity
        {
            get { return (IPersistEntity)GetValue(SelectedItemProperty); }
            set { SetValue(SelectedItemProperty, value); }
        }

        public static DependencyProperty SelectedItemProperty =
            DependencyProperty.Register("SelectedEntity", typeof(IPersistEntity), typeof(MainWindow), new FrameworkPropertyMetadata(null, FrameworkPropertyMetadataOptions.Inherits,
                                                                      OnPropertyChanged));
        // Model
        public IModel Model
        {
            get { return (IModel)GetValue(ModelProperty); }
            set { SetValue(ModelProperty, value); }
        }

        public static DependencyProperty ModelProperty =
            DependencyProperty.Register("Model", typeof(IModel), typeof(MainWindow), new FrameworkPropertyMetadata(null, FrameworkPropertyMetadataOptions.Inherits,
                                                                      OnPropertyChanged));
        
        #endregion

        #region UI Control

        public void BindUi(IXbimXplorerPluginMasterWindow mainWindow)
        {
            _xpWindow = mainWindow;

            SetBinding(SelectedItemProperty, new Binding("SelectedItem") { Source = mainWindow, Mode = BindingMode.TwoWay });
            SetBinding(SelectionProperty, new Binding("Selection") { Source = mainWindow.DrawingControl, Mode = BindingMode.TwoWay });
            SetBinding(ModelProperty, new Binding()); // whole datacontext binding, see http://stackoverflow.com/questions/8343928/how-can-i-create-a-binding-in-code-behind-that-doesnt-specify-a-path
        }

        private static void OnPropertyChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            // if any UI event should happen it needs to be specified here
            var window = d as MainWindow;
            if (window == null)
                return;

            switch (e.Property.Name)
            {
                case "Selection":
                    /*                    // Debug.WriteLine(e.Property.Name + @" changed");
                                        window.RefreshReport();*/
                    break;
                case "Model":
                    // Debug.WriteLine(e.Property.Name + @" changed");
                    /*window.WorkerEnsureStop();
                    if (window.Doc != null)
                    {
                        window.Doc.ClearCache();
                        if (window.AdaptSchema)
                        {
                            window.Doc.FixReferences();
                            window.UpdateUiLists();
                        }
                    }*/

                    window.GetElementList();
                    break;
                case "SelectedEntity":
                    /*window.RefreshReport();*/

                    break;
            }
        }

        public MainWindow()
        {
            InitializeComponent();

            string selectedDatabase = selectDatabase();

            if (selectedDatabase != null)
            {
                thread = new Thread(() => excuteCommandLine(selectedDatabase, openLCAPath));

                thread.Start();

                serverStarted = true;
            }

            string executableDirectory = AppDomain.CurrentDomain.BaseDirectory;

            string imagePath = Path.Combine(executableDirectory, "Resources", "openLCA.png");

            ImagePath.Source = new BitmapImage(new Uri(imagePath, UriKind.Absolute));

            //Subscribe to the loading model event of the main window

            //Also try to load model when open the plugin window
            GetElementList();
        }

        private void openLCA_Click(object sender, RoutedEventArgs e)
        {

            if (serverProcess != null)
            {
                serverProcess.KillTree();

                serverStarted = false;
            }

            string selectedDatabase = selectDatabase();

            if (selectedDatabase != null)
            {
                thread = new Thread(() => excuteCommandLine(selectedDatabase, openLCAPath));

                thread.Start();

                serverStarted = true;
            }
        }

        private string selectDatabase()
        {
            const string registryKey = @"SOFTWARE\openLCA";

            using (var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(registryKey))
            {
                if (key == null)
                {
                    MessageBox.Show("OpenLCA is not installed on your computer, server connection failed!");

                    return null;
                }

                openLCAPath = key.GetValue("Path").ToString();

                string openLCAExe = openLCAPath + "\\openLCA.exe";

                if (!File.Exists(openLCAExe))
                {
                    MessageBox.Show("Couldn't find the openLCA installation path, please make sure it is intalled");

                    return null;
                }

                //MessageBox.Show("openLCA is installed on your computer");

                string currentUserPath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);

                string openLCAWorkspace = currentUserPath + "\\openLCA-data-1.4";

                string libraryList = openLCAWorkspace + "\\databases.json";

                if (!File.Exists(libraryList))
                {
                    MessageBox.Show("openLCA workspace is not set up on your computer");

                    return null;
                }

                //MessageBox.Show(libraryList);

                try
                {
                    string jsonStrings = File.ReadAllText(libraryList);

                    Dictionary<string, List<string>> databaseDictionary = new Dictionary<string, List<string>>();

                    databaseDictionary.Add("Unsepecified", new List<string>());

                    // MessageBox.Show(libraryList);

                    var databases = Newtonsoft.Json.JsonConvert.DeserializeObject<Dictionary<string, Dictionary<string, string>[]>>(jsonStrings);

                    //MessageBox.Show("Step 1");

                    // Iterate over each database entry
                    foreach (var database in databases["localDatabases"])
                    {
                        string databaseName = database["name"];

                        // Check if the "category" property exists
                        if (database.ContainsKey("category"))
                        {
                            string databaseCategory = database["category"];

                            if (databaseDictionary.ContainsKey(databaseCategory))
                            {
                                databaseDictionary[databaseCategory].Add(databaseName);
                            }
                            else
                            {
                                List<string> databaseList = new List<string>();

                                databaseList.Add(databaseName);

                                databaseDictionary.Add(databaseCategory, databaseList);
                            }
                        }
                        else
                        {
                            databaseDictionary["Unsepecified"].Add(databaseName);
                        }
                    }

                    // MessageBox.Show("Database Dictionary Created");

                    if (databaseDictionary["Unsepecified"].Count > 0)
                    {
                        SelectDatabase newWindow = new SelectDatabase(databaseDictionary);

                        bool? dialogResult = newWindow.ShowDialog();

                        if (dialogResult.Value && dialogResult.HasValue)
                        {
                            return newWindow.databaseSelected;

                            //MessageBox.Show("You have selected " + selectedDatabase);                                
                        }
                        else
                        {
                            MessageBox.Show("You have not selected any database");
                        }
                    }
                    else
                    {
                        MessageBox.Show("Error: DatabaseList is empty");
                    }
                }
                catch (Exception e)
                {
                    MessageBox.Show("Error!");
                }
            }

            return null;
        }

        private void excuteCommandLine(string selectedDatabase, string openLCAPath)
        {
            string originalFilePath = openLCAPath + "\\bin\\grpc-server.cmd";

            string modifiedFilePath = openLCAPath + "\\bin\\grpc-server_Test.cmd";

            if (!File.Exists(modifiedFilePath))
            {
                try
                {
                    // Read all lines from the source file
                    string[] lines = File.ReadAllLines(originalFilePath);

                    // Search for specific words in each line
                    for (int i = 0; i < lines.Length; i++)
                    {
                        if (lines[i].Contains("-db %1"))
                        {
                            // Insert text after finding the specific words
                            lines[i] += " -p 8081";
                        }
                    }

                    // Write modified lines to the destination file
                    File.WriteAllLines(modifiedFilePath, lines);

                    MessageBox.Show("File copied, renamed, and modified successfully.");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("An error occurred: " + ex.Message);
                }
            }

            string commandLine = "cd " + openLCAPath + "\\bin";
            string commandLine2 = ".\\grpc-server_Test.cmd " + selectedDatabase;

            // Create a new process
            serverProcess = new Process();

            // Configure the process to run a command in a new console window
            ProcessStartInfo startInfo = new ProcessStartInfo();

            startInfo.FileName = "cmd.exe";
            startInfo.Arguments = "/k C: & " + commandLine + " & " + commandLine2;

            // Redirect standard output and error streams to the console
            startInfo.RedirectStandardOutput = true;
            startInfo.RedirectStandardError = true;
            startInfo.UseShellExecute = false;
            startInfo.CreateNoWindow = true;

            // Event handlers for asynchronous reading of output and error streams
            serverProcess.OutputDataReceived += (sender, e) => Console.WriteLine(e.Data);
            serverProcess.ErrorDataReceived += (sender, e) => Console.WriteLine(e.Data);

            // Start the process
            serverProcess = Process.Start(startInfo);

            // Begin asynchronous reading of output and error streams
            serverProcess.BeginOutputReadLine();
            serverProcess.BeginErrorReadLine();

            // Wait for the process to exit
            serverProcess.WaitForExit();
        }

        private void UserControl_Unloaded(object sender, RoutedEventArgs e)
        {
            if (serverProcess != null)
            {
                serverProcess.KillTree();
            }

            if (thread != null)
            {
                thread.Join();
            }
        }

        private void runButton_Click(object sender, RoutedEventArgs e)
        {

            //check if the model is loaded
            if (Model == null)
            {
                MessageBox.Show("Please load a model to run the calculation");

                return;
            }

            //check if the any stage is selected
            if (calculate_A1_A3.IsChecked == false && calculate_A4.IsChecked == false && calculate_A5.IsChecked == false)
            {
                MessageBox.Show("Please select at least one stage to calculate");

                return;
            }

            string mvd_description= Model.Header.ModelViewDefinition;

            // Regular expression pattern to match the values within [] followed by ViewDefinition
            string pattern = @"ViewDefinition\s*\[([^[\]]*)\]";

            // Match the pattern in the line
            Match match = Regex.Match(mvd_description, pattern);

            string mvd_Name ="";

            // Check if a match is found
            if (match.Success)
            {
                // Get the value enclosed within square brackets
                mvd_Name = match.Groups[1].Value;
            }

            if (mvd_Name != "BIM-LCA Integration View")
            {
                MessageBox.Show("Error! The target MVD was not BIM-LCA Integration View!, Please check you file and retry");

                return;
            }

            //check if the database is selected

            if (!serverStarted)
            {
                MessageBox.Show("Server Error! Please select open a openLCA database to start a server");

                return;
            }

            SelectedElements = new List<IfcElement>();

            UpdateSelectedList(elementList);

            //check if any element is selected
            if (SelectedElements.Count == 0)
            {
                MessageBox.Show("Please select at least one element to calculate");

                return;
            }

            List<string> eleInfo = GenerateRequiredInfo();

            if (eleInfo.Count == 0)
            {
                MessageBox.Show("No Success Info found");

                return;
            }

            string connectionString = $"Data Source=ae429-8914;Initial Catalog=MaterialDatabase;User Id=New User;Password=12345678;";

            if (!connect.ConnectDB(connectionString))
            {
                MessageBox.Show("Connect to SQL database failed, please retry!");

                return;
            }

            bool success_Size = int.TryParse(samplingSize.Text, out int samplingNum);

            if (!success_Size)
            {
                MessageBox.Show("Please enter a valid sampling number! 1-100 for sampling by percentage, an integer greater 0 for sampling by number");

                return;
            }

            int samplingMode = 0;

            if (byNum.IsChecked != true && byPct.IsChecked != true)
            {
                MessageBox.Show("Please specify a sampling method!");

                return;
            }
            else if (byNum.IsChecked == true)
            {
                if (samplingNum < 1)
                {
                    MessageBox.Show("Please enter a number greater than 1!");

                    return;
                }
                else
                {
                    samplingMode = 1;
                }
            }
            else if (byPct.IsChecked == true)
            {
                if (samplingNum < 1 || samplingNum > 100)
                {
                    MessageBox.Show("Please enter a number between 1 - 100!");

                    return;
                }
                else
                {
                    samplingMode = 2;
                }
            }

            List<List<MaterialInfo>> eleInfos = new List<List<MaterialInfo>>();

            DateTime timeStamp = DateTime.Now;

            foreach (var ele in eleInfo)
            {
                eleInfos.Add(RetrieveLCAInfo(ele, connect, samplingNum, samplingMode));
            }

            //PrintResult(eleInfos);

            connect.CleanUp(timeStamp);

            MessageBox.Show("Info Retrived!");

            CalculateLCA(eleInfos);

        }

        private void TextBox_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            // Check if the entered text is a valid integer
            if (int.TryParse(e.Text, out int num))
            {
                if (byPct.IsChecked == true && (num < 0 || num > 100))
                {
                    e.Handled = true;

                }
                else if (byNum.IsChecked == true && num < 0)
                {
                    e.Handled = true;
                }
            }
            else
            {
                // If not a valid integer, mark the event as handled
                e.Handled = true;
            }
        }

        private void byNum_Checked(object sender, RoutedEventArgs e)
        {
            if (byPct.IsChecked == true)
            {
                byPct.IsChecked = false;
            }
        }

        private void byPct_Checked(object sender, RoutedEventArgs e)
        {
            if (byNum.IsChecked == true)
            {
                byNum.IsChecked = false;
            }
        }

        private void GetElementList()
        {
            if (Model == null)
            {
                return;
            }
            //get all the elements in the model

            elementList = GetList();

            itemForSelection.ItemsSource = elementList;

        }

        private List<TreeNode<string>> GetList()
        {
            List<TreeNode<string>> elementList = new List<TreeNode<string>>();
            //Wall, Door, Window, Slab, Roof, Stair, Ramp, Beam, Column, Covering, Foundation, Pile
            TreeNode<string> mainCompoenents = new TreeNode<string>("Main Compoenets");

            var doors = Model.Instances.OfType<IfcDoor>().ToList();

            if (doors.Count > 0)
            {
                TreeNode<string> Doors = new TreeNode<string>("Doors");

                foreach (var door in doors)
                {
                    TreeNode<string> doorNode = new TreeNode<string>(door.Name);

                    doorNode.Element = door;

                    Doors.Children.Add(doorNode);

                    doorNode.ParentNode = Doors;
                }

                mainCompoenents.Children.Add(Doors);

                Doors.ParentNode = mainCompoenents;
            }

            var windows = Model.Instances.OfType<IfcWindow>().ToList();

            if (windows.Count > 0)
            {
                TreeNode<string> Windows = new TreeNode<string>("Windows");

                foreach (var window in windows)
                {
                    TreeNode<string> windowNode = new TreeNode<string>(window.Name);

                    windowNode.Element = window;

                    Windows.Children.Add(windowNode);

                    windowNode.ParentNode = Windows;
                }

                mainCompoenents.Children.Add(Windows);

                Windows.ParentNode = mainCompoenents;
            }

            var walls = Model.Instances.OfType<IfcWall>().ToList();

            TreeNode<string> Walls = new TreeNode<string>("Walls");

            if (walls.Count > 0)
            {
                foreach (var wall in walls)
                {
                    TreeNode<string> wallNode = new TreeNode<string>(wall.Name);

                    wallNode.Element = wall;

                    Walls.Children.Add(wallNode);

                    wallNode.ParentNode = Walls;
                }
            }

            var curtainWalls = Model.Instances.OfType<IfcCurtainWall>().ToList();

            if (curtainWalls.Count > 0)
            {
                TreeNode<string> CurtainWalls = new TreeNode<string>("Curtain Walls");

                foreach (var curtainWall in curtainWalls)
                {
                    TreeNode<string> curtainWallNode = new TreeNode<string>(curtainWall.Name);

                    curtainWallNode.Element = curtainWall;

                    CurtainWalls.Children.Add(curtainWallNode);

                    curtainWallNode.ParentNode = CurtainWalls;
                }

                Walls.Children.Add(CurtainWalls);

                CurtainWalls.ParentNode = Walls;
            }

            if (Walls.Children.Count > 0)
            {
                mainCompoenents.Children.Add(Walls);

                Walls.ParentNode = mainCompoenents;
            }

            var floors = Model.Instances.OfType<IfcSlab>().ToList();

            if (floors.Count > 0)
            {
                TreeNode<string> Floors = new TreeNode<string>("Floors");

                foreach (var floor in floors)
                {
                    TreeNode<string> floorNode = new TreeNode<string>(floor.Name);

                    floorNode.Element = floor;

                    Floors.Children.Add(floorNode);

                    floorNode.ParentNode = Floors;
                }

                mainCompoenents.Children.Add(Floors);

                Floors.ParentNode = mainCompoenents;
            }

            var roofs = Model.Instances.OfType<IfcRoof>().ToList();

            if (roofs.Count > 0)
            {
                TreeNode<string> Roofs = new TreeNode<string>("Roofs");

                foreach (var roof in roofs)
                {
                    TreeNode<string> roofNode = new TreeNode<string>(roof.Name);

                    roofNode.Element = roof;

                    Roofs.Children.Add(roofNode);

                    roofNode.ParentNode = Roofs;
                }

                mainCompoenents.Children.Add(Roofs);

                Roofs.ParentNode = mainCompoenents;
            }

            var ceilings = Model.Instances.OfType<IfcCovering>().ToList();

            if (ceilings.Count > 0)
            {
                TreeNode<string> Ceilings = new TreeNode<string>("Ceilings");

                foreach (var covering in ceilings)
                {
                    TreeNode<string> ceilingNode = new TreeNode<string>(covering.Name);

                    ceilingNode.Element = covering;

                    Ceilings.Children.Add(ceilingNode);

                    ceilingNode.ParentNode = Ceilings;
                }

                mainCompoenents.Children.Add(Ceilings);

                Ceilings.ParentNode = mainCompoenents;
            }

            var stairs = Model.Instances.OfType<IfcStair>().ToList();

            if (stairs.Count > 0)
            {
                TreeNode<string> Stairs = new TreeNode<string>("Stairs");

                foreach (var stair in stairs)
                {
                    TreeNode<string> stairNode = new TreeNode<string>(stair.Name);

                    stairNode.Element = stair;

                    var relatingObjects = stair.IsDecomposedBy.First().RelatedObjects;

                    foreach (var subEle in relatingObjects)
                    {
                        if (subEle is IfcStairFlight)
                        {
                            TreeNode<string> stairFlightNode = new TreeNode<string>(subEle.Name);

                            stairFlightNode.Element = subEle as IfcStairFlight;

                            stairNode.Children.Add(stairFlightNode);

                            stairFlightNode.ParentNode = stairNode;
                        }
                        else if (subEle is IfcSlab)
                        {
                            TreeNode<string> landingNode = new TreeNode<string>(subEle.Name);

                            landingNode.Element = subEle as IfcSlab;

                            stairNode.Children.Add(landingNode);

                            landingNode.ParentNode = stairNode;
                        }
                        else if (subEle is IfcRailing)
                        {
                            TreeNode<string> railingNode = new TreeNode<string>(subEle.Name);

                            railingNode.Element = subEle as IfcRailing;

                            stairNode.Children.Add(railingNode);

                            railingNode.ParentNode = stairNode;
                        }
                    }

                    Stairs.Children.Add(stairNode);

                    stairNode.ParentNode = Stairs;
                }

                mainCompoenents.Children.Add(Stairs);

                Stairs.ParentNode = mainCompoenents;
            }

            var ramps = Model.Instances.OfType<IfcRamp>().ToList();

            if (ramps.Count > 0)
            {
                TreeNode<string> Ramps = new TreeNode<string>("Ramps");

                foreach (var ramp in ramps)
                {
                    TreeNode<string> rampNode = new TreeNode<string>(ramp.Name);

                    rampNode.Element = ramp;

                    var relatingObjects = ramp.IsDecomposedBy.First().RelatedObjects;

                    foreach (var subEle in relatingObjects)
                    {
                        if (subEle is IfcRampFlight)
                        {
                            TreeNode<string> rampFlightNode = new TreeNode<string>(subEle.Name);

                            rampFlightNode.Element = subEle as IfcRampFlight;

                            rampNode.Children.Add(rampFlightNode);

                            rampFlightNode.ParentNode = rampNode;
                        }
                        else if (subEle is IfcSlab)
                        {
                            TreeNode<string> landingNode = new TreeNode<string>(subEle.Name);

                            landingNode.Element = subEle as IfcSlab;

                            rampNode.Children.Add(landingNode);

                            landingNode.ParentNode = rampNode;
                        }
                        else if (subEle is IfcRailing)
                        {
                            TreeNode<string> railingNode = new TreeNode<string>(subEle.Name);

                            railingNode.Element = subEle as IfcRailing;

                            rampNode.Children.Add(railingNode);

                            railingNode.ParentNode = rampNode;
                        }
                    }

                    Ramps.Children.Add(rampNode);

                    rampNode.ParentNode = Ramps;
                }

                mainCompoenents.Children.Add(Ramps);

                Ramps.ParentNode = mainCompoenents;
            }

            var beams = Model.Instances.OfType<IfcBeam>().ToList();

            if (beams.Count > 0)
            {
                TreeNode<string> Beams = new TreeNode<string>("Beams");

                foreach (var beam in beams)
                {
                    TreeNode<string> beamNode = new TreeNode<string>(beam.Name);

                    beamNode.Element = beam;

                    Beams.Children.Add(beamNode);

                    beamNode.ParentNode = Beams;
                }

                mainCompoenents.Children.Add(Beams);

                Beams.ParentNode = mainCompoenents;
            }

            var columns = Model.Instances.OfType<IfcColumn>().ToList();

            if (columns.Count > 0)
            {
                TreeNode<string> Columns = new TreeNode<string>("Columns");

                foreach (var column in columns)
                {
                    TreeNode<string> columnNode = new TreeNode<string>(column.Name);

                    columnNode.Element = column;

                    Columns.Children.Add(columnNode);

                    columnNode.ParentNode = Columns;
                }

                mainCompoenents.Children.Add(Columns);

                Columns.ParentNode = mainCompoenents;
            }

            TreeNode<string> Foundations = new TreeNode<string>("Foundations");

            var footings = Model.Instances.OfType<IfcFooting>().ToList();

            if (footings.Count > 0)
            {
                TreeNode<string> Footings = new TreeNode<string>("Footings");

                foreach (var footing in footings)
                {
                    TreeNode<string> footingNode = new TreeNode<string>(footing.Name);

                    footingNode.Element = footing;

                    Footings.Children.Add(footingNode);

                    footingNode.ParentNode = Foundations;
                }

                Foundations.Children.Add(Footings);

                Footings.ParentNode = Foundations;
            }

            var piles = Model.Instances.OfType<IfcPile>().ToList();

            if (piles.Count > 0)
            {
                TreeNode<string> Piles = new TreeNode<string>("Piles");

                foreach (var p in piles)
                {
                    TreeNode<string> pileNode = new TreeNode<string>(p.Name);
                    pileNode.Element = p;
                    Piles.Children.Add(pileNode);
                    pileNode.ParentNode = Piles;
                }

                Foundations.Children.Add(Piles);

                Piles.ParentNode = Foundations;
            }

            if (Foundations.Children.Count > 0)
            {
                mainCompoenents.Children.Add(Foundations);

                Foundations.ParentNode = mainCompoenents;
            }

            var chimenies = Model.Instances.OfType<IfcChimney>().ToList();

            if (chimenies.Count > 0)
            {
                TreeNode<string> Chimeny = new TreeNode<string>("Chimeny");

                foreach (var chimeny in chimenies)
                {
                    TreeNode<string> chimenyNode = new TreeNode<string>(chimeny.Name);

                    chimenyNode.Element = chimeny;

                    Chimeny.Children.Add(chimenyNode);

                    chimenyNode.ParentNode = Chimeny;
                }

                mainCompoenents.Children.Add(Chimeny);

                Chimeny.ParentNode = mainCompoenents;
            }

            TreeNode<string> Others = new TreeNode<string>("Others");

            var proxies = Model.Instances.OfType<IfcBuildingElementProxy>().ToList();

            if (proxies.Count > 0)
            {
                foreach (var other in proxies)
                {
                    TreeNode<string> otherNode = new TreeNode<string>(other.Name);

                    otherNode.Element = other;

                    Others.Children.Add(otherNode);

                    otherNode.ParentNode = Others;
                }
            }

            var members = Model.Instances.OfType<IfcMember>().ToList();

            if (members.Count > 0)
            {
                foreach (var other in members)
                {
                    TreeNode<string> otherNode = new TreeNode<string>(other.Name);

                    otherNode.Element = other;

                    Others.Children.Add(otherNode);

                    otherNode.ParentNode = Others;
                }
            }

            var plates = Model.Instances.OfType<IfcPlate>().ToList();

            if (plates.Count > 0)
            {
                foreach (var other in plates)
                {
                    TreeNode<string> otherNode = new TreeNode<string>(other.Name);

                    otherNode.Element = other;

                    Others.Children.Add(otherNode);

                    otherNode.ParentNode = Others;
                }
            }

            if (Others.Children.Count > 0)
            {
                mainCompoenents.Children.Add(Others);

                Others.ParentNode = mainCompoenents;
            }

            TreeNode<string> Reinforcement = new TreeNode<string>("Reinforcement");

            var reinforcingBar = Model.Instances.OfType<IfcReinforcingBar>().ToList();

            if(reinforcingBar.Count>0)
            {
                TreeNode<string> reinforcement_Bar = new TreeNode<string>("Reinforcing Bars");

                foreach (var bar in reinforcingBar)
                {
                    TreeNode<string> barNode = new TreeNode<string>(bar.Name);

                    barNode.Element = bar;

                    reinforcement_Bar.Children.Add(barNode);

                    barNode.ParentNode = reinforcement_Bar;
                }

                Reinforcement.Children.Add(reinforcement_Bar);

                reinforcement_Bar.ParentNode = Reinforcement;
            }

            var reinforcingMesh = Model.Instances.OfType<IfcReinforcingMesh>().ToList();

            if(reinforcingMesh.Count>0)
            {
                TreeNode<string> reinforcement_Mesh = new TreeNode<string>("Reinforcing Meshes");

                foreach (var mesh in reinforcingMesh)
                {
                    TreeNode<string> meshNode = new TreeNode<string>(mesh.Name);

                    meshNode.Element = mesh;

                    reinforcement_Mesh.Children.Add(meshNode);

                    meshNode.ParentNode = reinforcement_Mesh;
                }

                Reinforcement.Children.Add(reinforcement_Mesh);

                reinforcement_Mesh.ParentNode = Reinforcement;
            }

            if (Reinforcement.Children.Count > 0)
            {
                mainCompoenents.Children.Add(Reinforcement);

                Reinforcement.ParentNode = mainCompoenents;
            }

            if (mainCompoenents.Children.Count > 0)
            {
                elementList.Add(mainCompoenents);
            }         

            TreeNode<string> MEPSystems = new TreeNode<string>("MEP Systems");

            var plumbingFixtures = Model.Instances.OfType<IfcSanitaryTerminal>().ToList();

            if (plumbingFixtures.Count > 0)
            {
                TreeNode<string> PlumbingFixtures = new TreeNode<string>("Plumbing Fixtures");

                foreach (var plumbingFixture in plumbingFixtures)
                {
                    TreeNode<string> plumbingFixtureNode = new TreeNode<string>(plumbingFixture.Name);

                    plumbingFixtureNode.Element = plumbingFixture;

                    PlumbingFixtures.Children.Add(plumbingFixtureNode);

                    plumbingFixtureNode.ParentNode = PlumbingFixtures;
                }

                MEPSystems.Children.Add(PlumbingFixtures);

                PlumbingFixtures.ParentNode = MEPSystems;
            }

            TreeNode<string> LightingFixtures = new TreeNode<string>("Lighting Fixtures");

            var lightingFixtures = Model.Instances.OfType<IfcLightFixture>().ToList();

            if (lightingFixtures.Count > 0)
            {
                foreach (var lightingFixture in lightingFixtures)
                {
                    TreeNode<string> lightingFixtureNode = new TreeNode<string>(lightingFixture.Name);

                    lightingFixtureNode.Element = lightingFixture;

                    LightingFixtures.Children.Add(lightingFixtureNode);

                    lightingFixtureNode.ParentNode = LightingFixtures;
                }
            }

            var lamps = Model.Instances.OfType<IfcLamp>().ToList();

            if (lamps.Count > 0)
            {
                foreach (var lamp in lamps)
                {
                    TreeNode<string> lampNode = new TreeNode<string>(lamp.Name);

                    lampNode.Element = lamp;

                    LightingFixtures.Children.Add(lampNode);

                    lampNode.ParentNode = LightingFixtures;
                }
            }

            if (LightingFixtures.Children.Count > 0)
            {
                MEPSystems.Children.Add(LightingFixtures);

                LightingFixtures.ParentNode = MEPSystems;
            }

            var controls = Model.Instances.OfType<IfcController>().ToList();

            if (controls.Count > 0)
            {
                TreeNode<string> Controls = new TreeNode<string>("Controls");

                foreach (var control in controls)
                {
                    TreeNode<string> controlNode = new TreeNode<string>(control.Name);

                    controlNode.Element = control;

                    Controls.Children.Add(controlNode);

                    controlNode.ParentNode = Controls;
                }

                MEPSystems.Children.Add(Controls);

                Controls.ParentNode = MEPSystems;
            }

            TreeNode<string> Sensors = new TreeNode<string>("Sensors");

            var sensors = Model.Instances.OfType<IfcSensor>().ToList();

            if (sensors.Count > 0)
            {
                foreach (var sensor in sensors)
                {
                    TreeNode<string> sensorNode = new TreeNode<string>(sensor.Name);

                    sensorNode.Element = sensor;

                    Sensors.Children.Add(sensorNode);

                    sensorNode.ParentNode = Sensors;
                }

                MEPSystems.Children.Add(Sensors);

                Sensors.ParentNode = MEPSystems;
            }

            if (MEPSystems.Children.Count > 0)
            {
                elementList.Add(MEPSystems);
            }

            TreeNode<string> HVAC = new TreeNode<string>("HVAC");

            var airTerminals = Model.Instances.OfType<IfcAirTerminal>().ToList();

            if (airTerminals.Count > 0)
            {
                TreeNode<string> AirTerminals = new TreeNode<string>("Air Terminals");

                foreach (var airTerminal in airTerminals)
                {
                    TreeNode<string> airTerminalNode = new TreeNode<string>(airTerminal.Name);

                    airTerminalNode.Element = airTerminal;

                    AirTerminals.Children.Add(airTerminalNode);

                    airTerminalNode.ParentNode = AirTerminals;
                }

                HVAC.Children.Add(AirTerminals);

                AirTerminals.ParentNode = HVAC;
            }

            var ducts = Model.Instances.OfType<IfcDuctSegment>().ToList();

            if (ducts.Count > 0)
            {
                TreeNode<string> Ducts = new TreeNode<string>("Ducts");

                foreach (var duct in ducts)
                {
                    TreeNode<string> ductNode = new TreeNode<string>(duct.Name);

                    ductNode.Element = duct;

                    Ducts.Children.Add(ductNode);

                    ductNode.ParentNode = Ducts;
                }

                HVAC.Children.Add(Ducts);

                Ducts.ParentNode = HVAC;
            }

            var fans = Model.Instances.OfType<IfcFan>().ToList();

            if (fans.Count > 0)
            {
                TreeNode<string> Fans = new TreeNode<string>("Fans");

                foreach (var fan in fans)
                {
                    TreeNode<string> fanNode = new TreeNode<string>(fan.Name);

                    fanNode.Element = fan;

                    Fans.Children.Add(fanNode);

                    fanNode.ParentNode = Fans;
                }

                HVAC.Children.Add(Fans);

                Fans.ParentNode = HVAC;
            }

            var pumps = Model.Instances.OfType<IfcPump>().ToList();

            if (pumps.Count > 0)
            {
                TreeNode<string> Pumps = new TreeNode<string>("Pumps");

                foreach (var pump in pumps)
                {
                    TreeNode<string> pumpNode = new TreeNode<string>(pump.Name);

                    pumpNode.Element = pump;

                    Pumps.Children.Add(pumpNode);

                    pumpNode.ParentNode = Pumps;
                }

                HVAC.Children.Add(Pumps);

                Pumps.ParentNode = HVAC;

            }

            var heatTanks = Model.Instances.OfType<IfcTank>().ToList();

            if (heatTanks.Count > 0)
            {
                TreeNode<string> HeatTanks = new TreeNode<string>("Tanks");

                foreach (var heatTank in heatTanks)
                {
                    TreeNode<string> heatTankNode = new TreeNode<string>(heatTank.Name);

                    heatTankNode.Element = heatTank;

                    HeatTanks.Children.Add(heatTankNode);

                    heatTankNode.ParentNode = HeatTanks;
                }

                HVAC.Children.Add(HeatTanks);

                HeatTanks.ParentNode = HVAC;
            }

            if (HVAC.Children.Count > 0)
            {
                elementList.Add(HVAC);
            }

            var furnishing = Model.Instances.OfType<IfcFurnishingElement>().ToList();

            if (furnishing.Count > 0)
            {
                TreeNode<string> Furnishing = new TreeNode<string>("Furnishing");

                foreach (var f in furnishing)
                {
                    TreeNode<string> furnishingNode = new TreeNode<string>(f.Name);

                    furnishingNode.Element = f;

                    Furnishing.Children.Add(furnishingNode);

                    furnishingNode.ParentNode = Furnishing;
                }

                elementList.Add(Furnishing);
            }

            return GetSortedList(elementList);

        }

        private List<TreeNode<string>> GetSortedList(List<TreeNode<string>> elementList)
        {
            foreach (var element in elementList)
            {
                if (element.Children.Count == 0)
                {
                    continue;
                }

                element.Children = GetSortedList(element.Children);
            }

            return elementList.OrderBy(x => x.Data).ToList();
        }

        private List<string> GenerateRequiredInfo()
        {
            List<string> successInstances = new List<string>();

            List<string> failedInstances = new List<string>();

            List<string> distinctedList = SelectedElements.Select(x => x.EntityLabel.ToString()).Distinct().ToList();

            foreach (var element in SelectedElements)
            {
                var label = element.EntityLabel.ToString();

                var name = element.Name.ToString();

                var ifctype_name = element.GetType().Name;

                string type_name = "";

                if (element is IfcReinforcingBar )
                {
                   var reference=element.PropertySets.Where(x => x.Name.Value.ToString() == "Pset_ReinforcingBarCommon").FirstOrDefault().HasProperties.Where(x => x.Name.Value.ToString() == "Reference").FirstOrDefault() as IfcPropertySingleValue;

                    if(reference!=null)
                    {
                        type_name = reference.NominalValue.ToString();
                    }
                    else
                    {
                        type_name = "Unspecified";
                    }

                    string[] pair = { label, name, ifctype_name, type_name, "0710f1c2-7c3b-11ee-b962-0242ac120002-XXX-7850" };

                    successInstances.Add(string.Join("\t", pair));
                }
                else if(element is IfcReinforcingMesh)
                {
                    var reference = element.PropertySets.Where(x => x.Name.Value.ToString() == "Pset_ReinforcingMeshCommon").FirstOrDefault().HasProperties.Where(x => x.Name.Value.ToString() == "Reference").FirstOrDefault() as IfcPropertySingleValue;

                    if (reference != null)
                    {
                        type_name = reference.NominalValue.ToString();
                    }
                    else
                    {
                        type_name = "Unspecified";
                    }

                    string[] pair = { label, name, ifctype_name, type_name, "0710f1c2-7c3b-11ee-b962-0242ac120002-XXX-7850" };

                    successInstances.Add(string.Join("\t", pair));
                }
                else
                {
                    var type = element.IsTypedBy.FirstOrDefault().RelatingType;

                    if (type == null)
                    {
                        continue;
                    }
                   
                    type_name = type.Name.ToString();

                    var type_propertyset = type.HasPropertySets.Where(x => x.Name.Value.ToString() == "Cpset_MaterialInformation").FirstOrDefault() as IfcPropertySet;

                    string ID_value = "";

                    if (type_propertyset != null)
                    {
                        var ID = type_propertyset.HasProperties.Where(x => x.Name.Value.ToString() == "ID").FirstOrDefault() as IfcPropertySingleValue;

                        ID_value = ID.NominalValue.ToString();

                        string[] pair = { label, name, ifctype_name, type_name, ID_value };

                        successInstances.Add(string.Join("\t", pair));
                    }
                    else
                    {
                        string[] pair = { label, name, ifctype_name, type_name, ID_value };

                        failedInstances.Add(string.Join("\t", pair));
                    }
                }
            }

            //Write failed instances to a csv file

            WriteInstancesToCSV(successInstances,"Success Instances");

            WriteInstancesToCSV(failedInstances, "Failed Instances");

            //WriteInstancesToCSV(failedInstances);

            return successInstances;
        }

        private void UpdateSelectedList(List<TreeNode<string>> inputList)
        {
            foreach (var element in inputList)
            {
                if (element.IsSelected != false)
                {
                    if (element.Element != null && element.Children.Count == 0)
                    {
                        SelectedElements.Add(element.Element);
                    }
                    else if (element.Children.Count > 0)
                    {
                        UpdateSelectedList(element.Children);
                    }
                }
            }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            Button button = sender as Button;

            if (button == null)
            {
                return;
            }

            Brush color = new SolidColorBrush(Colors.White);

            Brush defaultColor = new SolidColorBrush(Colors.Gray);

            switch (button.Name)
            {
                case "A1_A3":
                    A1_A3.Background = color;
                    A4.Background = defaultColor;
                    A5.Background = defaultColor;
                    total.Background = defaultColor;
                    all_Stages.Background = defaultColor;
                    status_1 = "A1-A3";
                    break;
                case "A4":
                    A1_A3.Background = defaultColor;
                    A4.Background = color;
                    A5.Background = defaultColor;
                    total.Background = defaultColor;
                    all_Stages.Background = defaultColor;
                    status_1 = "A4";
                    break;
                case "A5":
                    A1_A3.Background = defaultColor;
                    A4.Background = defaultColor;
                    A5.Background = color;
                    total.Background = defaultColor;
                    all_Stages.Background = defaultColor;
                    status_1 = "A5";
                    break;
                case "total":
                    A1_A3.Background = defaultColor;
                    A4.Background = defaultColor;
                    A5.Background = defaultColor;
                    total.Background = color;
                    all_Stages.Background = defaultColor;
                    status_1 = "Total";
                    break;
                case "all_Stages":
                    A1_A3.Background = defaultColor;
                    A4.Background = defaultColor;
                    A5.Background = defaultColor;
                    total.Background = defaultColor;
                    all_Stages.Background = color;
                    status_1 = "All Stages";
                    break;
                case "min":
                    min.Background = color;
                    max.Background = defaultColor;
                    Avg.Background = defaultColor;
                    SD.Background = defaultColor;
                    status_2 = "Min";
                    break;
                case "max":
                    min.Background = defaultColor;
                    max.Background = color;
                    Avg.Background = defaultColor;
                    SD.Background = defaultColor;
                    status_2 = "Max";
                    break;
                case "Avg":
                    min.Background = defaultColor;
                    max.Background = defaultColor;
                    Avg.Background = color;
                    SD.Background = defaultColor;
                    status_2 = "Avg";
                    break;
                case "SD":
                    SD.Background = color;
                    min.Background = defaultColor;
                    max.Background = defaultColor;
                    Avg.Background = defaultColor;                 
                    status_2 = "SD";
                    break;
            }

            if (status_1 == null || status_2 == null)
            {
                return;
            }
            else if (status_1 == "A1_A3" && status_2 == "Avg")
            {
                resultTable.ItemsSource = A1_A3_Avg;
            }
            else if (status_1 == "A4" && status_2 == "Avg")
            {
                resultTable.ItemsSource = A4_Avg;
            }
            else if (status_1 == "A5" && status_2 == "Avg")
            {
                resultTable.ItemsSource = A5_Avg;
            }
            else if (status_1 == "Total" && status_2 == "Avg")
            {
                resultTable.ItemsSource = Total_Avg;
            }
            else if (status_1 == "All Stages" && status_2 == "Avg")
            {
                resultTable.ItemsSource = All_Stage_Avg;
            }
            else if (status_1 == "A1_A3" && status_2 == "Min")
            {
                resultTable.ItemsSource = A1_A3_Min;
            }
            else if (status_1 == "A4" && status_2 == "Min")
            {
                resultTable.ItemsSource = A4_Min;
            }
            else if (status_1 == "A5" && status_2 == "Min")
            {
                resultTable.ItemsSource = A5_Min;
            }
            else if (status_1 == "Total" && status_2 == "Min")
            {
                resultTable.ItemsSource = Total_Min;
            }
            else if (status_1 == "All Stages" && status_2 == "Min")
            {
                resultTable.ItemsSource = All_Stage_Min;
            }
            else if (status_1 == "A1_A3" && status_2 == "Max")
            {
                resultTable.ItemsSource = A1_A3_Max;
            }
            else if (status_1 == "A4" && status_2 == "Max")
            {
                resultTable.ItemsSource = A4_Max;
            }
            else if (status_1 == "A5" && status_2 == "Max")
            {
                resultTable.ItemsSource = A5_Max;
            }
            else if (status_1 == "Total" && status_2 == "Max")
            {
                resultTable.ItemsSource = Total_Max;
            }
            else if (status_1 == "All Stages" && status_2 == "Max")
            {
                resultTable.ItemsSource = All_Stage_Max;
            }
            else if (status_1=="A1-A3" && status_2=="SD")
            {
                resultTable.ItemsSource = A1_A3_SD;
            }
            else if (status_1 == "A4" && status_2 == "SD")
            {
                resultTable.ItemsSource = A4_SD;
            }
            else if (status_1 == "A5" && status_2 == "SD")
            {
                resultTable.ItemsSource = A5_SD;
            }
            else if (status_1 == "Total" && status_2 == "SD")
            {
                resultTable.ItemsSource = Total_SD;
            }
            else if (status_1 == "All Stages" && status_2 == "SD")
            {
                resultTable.ItemsSource = All_Stage_SD;
            }
        }

        private void reportButton_Click(object sender, RoutedEventArgs e)
        {
            //Open desktop folder
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

            Process.Start("explorer.exe", desktopPath);
        }

        #endregion

        #region Test Methods

        private void PrintLCARecordsResult(List<List<MaterialInfo>> eleInfos)
        {

            string filePath = "output.csv";

            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

            string fullPath = Path.Combine(desktopPath, filePath);

            int eleCount = 1;

            using (var writer = new StreamWriter(fullPath))
            {
                foreach (var ele in eleInfos)
                {
                    int matCount = 1;

                    if (ele.Count == 0)
                    {
                        writer.WriteLine($"Element {eleCount} / {eleInfos.Count}, No Material Found");

                        continue;
                    }

                    foreach (var material in ele)
                    {

                        if (material.MaterialRecords.Count == 0)
                        {
                            writer.WriteLine($"Element {eleCount} / {eleInfos.Count}, Material {eleCount} - {matCount}, {material.IfcLabel}, {material.Name}, {material.IfcType}, {material.ElementType}, {material.ID}, No LCA Data Found");

                            continue;
                        }

                        foreach (var record in material.MaterialRecords)
                        {
                            writer.WriteLine($"Element {eleCount} / {eleInfos.Count}, Material {eleCount} - {matCount}, {material.IfcLabel}, {material.Name}, {material.IfcType}, {material.ElementType}, {material.ID}, " +
                                $"{record.Uid}, {record.Unit}, {record.GWP}, {record.ODP}, {record.AP}, {record.EP}, {record.POCP},{record.ADPE},{record.ADPF},{record.PERT},{record.PENRT},{record.Address}");
                        }

                        matCount++;
                    }

                    eleCount++;
                }
            }

            MessageBox.Show("The result has been saved to the desktop as output.csv");
        }

        private void WriteInstancesToCSV(List<string> Instances,string fileName)
        {
            // write the database info to a csv file

            string desktopFolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

            string filePath = Path.Combine(desktopFolder, $"{fileName}.csv");

            int i = 1;

            while (File.Exists(filePath))
            {
                filePath = Path.Combine(desktopFolder, $"{fileName} ({i}).csv");

                i++;
            }

            using (var writer = new StreamWriter(filePath))
            {
                // Iterate over each string in the list
                foreach (string str in Instances)
                {
                    // Write the string followed by a comma to separate values
                    writer.WriteLine(str);
                }
            }
        }

        private void PrintQuantities(List<List<MaterialInfo>> InputData)
        {

            string desktopFolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

            string filePath = Path.Combine(desktopFolder, "Quantities.csv");

            int i = 1;

            while (File.Exists(filePath))
            {
                filePath = Path.Combine(desktopFolder, $"Quantities ({i}).csv");

                i++;
            }

            using (var writer = new StreamWriter(filePath))
            {
                // Iterate over each string in the list
                int eleCount = 1;

                foreach (var ele in InputData)
                {
                    int matCount = 1;

                    foreach (var material in ele)
                    {
                        writer.WriteLine($"Element {eleCount} / {InputData.Count}\t{material.IfcLabel}\t{material.IfcType}\t{material.ElementType}\tMaterial {matCount} - {material.Name}\t{material.ID}\t{material.quantity}\t{material.unit}\t{material.MaterialRecords.Count}");

                        matCount++;
                    }

                    eleCount++;
                }
            }

        }

        private void WriteLCAResult(List<EleLCAResult> results,string phase)
        {
            string desktopFolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

            string filePath = Path.Combine(desktopFolder, $"{phase} Results.txt");

            int i = 1;

            while (File.Exists(filePath))
            {
                filePath = Path.Combine(desktopFolder, $"{phase} Results ({i}).txt");

                i++;
            }

            using (var writer = new StreamWriter(filePath))
            {
                // Iterate over each string in the list
                int eleCount = 1;

                foreach (var ele in results)
                {
                    writer.WriteLine($"Element {eleCount}\t{ele.IfcLabel}\t{ele.Name}\t{ele.Stage}\t{ele.Description}\t{ele.MaterialResults.Count}/{ele.materialCount}\t{ele.TotalResult_Average.GWP}\t{ele.TotalResult_Average.ODP}\t{ele.TotalResult_Average.AP}\t{ele.TotalResult_Average.EP}\t{ele.TotalResult_Average.POCP}\t{ele.TotalResult_Average.ADPE}\t{ele.TotalResult_Average.ADPF}\t{ele.TotalResult_Average.PENRT}\t{ele.TotalResult_Average.PERT}");

                    eleCount++;
                }
            }
        }

        private EleLCAResult WriteResult(List<EleLCAResult> results, string phase)
        {
            string desktopFolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

            string filePath = Path.Combine(desktopFolder, $"{phase} Results.txt");

            int i = 1;

            while (File.Exists(filePath))
            {
                filePath = Path.Combine(desktopFolder, $"{phase} Results ({i}).txt");

                i++;
            }

            Record totalRecord_avg = new Record()
            {
                GWP = 0,
                ODP = 0,
                AP = 0,
                EP = 0,
                POCP = 0,
                ADPE = 0,
                ADPF = 0,
                PERT = 0,
                PENRT = 0
            };

            Record totalRecord_min = new Record()
            {
                GWP = 0,
                ODP = 0,
                AP = 0,
                EP = 0,
                POCP = 0,
                ADPE = 0,
                ADPF = 0,
                PERT = 0,
                PENRT = 0
            };

            Record totalRecord_max = new Record()
            {
                GWP = 0,
                ODP = 0,
                AP = 0,
                EP = 0,
                POCP = 0,
                ADPE = 0,
                ADPF = 0,
                PERT = 0,
                PENRT = 0           
            };

            Record totalRecord_sd = new Record()
            {
                GWP = 0,
                ODP = 0,
                AP = 0,
                EP = 0,
                POCP = 0,
                ADPE = 0,
                ADPF = 0,
                PERT = 0,
                PENRT = 0
            };

            using (var writer = new StreamWriter(filePath))
            {
                // Iterate over each string in the list
                int eleCount = 1;

                writer.WriteLine($"{phase}\tGWP\t\t\t\tODP\t\t\t\tAP\t\t\t\tEP\t\t\t\tPOCP\t\t\t\tADPE\t\t\t\tADPF\t\t\t\tPERT\t\t\t\tPENRT\t\t\t\t");
                writer.WriteLine("Element No.\tMin\tMax\tAvg\tSD\tMin\tMax\tAvg\tSD\tMin\tMax\tAvg\tSD\tMin\tMax\tAvg\tSD\tMin\tMax\tAvg\tSD\tMin\tMax\tAvg\tSD\tMin\tMax\tAvg\tSD\tMin\tMax\tAvg\tSD\tMin\tMax\tAvg\tSD");
               
                foreach (var ele in results)
                {

                    writer.WriteLine($"No.{eleCount}: {ele.Name}\t{ele.TotalResult_Min.GWP}\t{ele.TotalResult_Max.GWP}\t{ele.TotalResult_Average.GWP}\t{ele.TotalResult_SD.GWP}\t{ele.TotalResult_Min.ODP}\t{ele.TotalResult_Max.ODP}\t{ele.TotalResult_Average.ODP}\t{ele.TotalResult_SD.ODP}\t{ele.TotalResult_Min.AP}\t{ele.TotalResult_Max.AP}\t{ele.TotalResult_Average.AP}\t{ele.TotalResult_SD.AP}\t{ele.TotalResult_Min.EP}\t{ele.TotalResult_Max.EP}\t{ele.TotalResult_Average.EP}\t{ele.TotalResult_SD.EP}\t{ele.TotalResult_Min.POCP}\t{ele.TotalResult_Max.POCP}\t{ele.TotalResult_Average.POCP}\t{ele.TotalResult_SD.POCP}\t{ele.TotalResult_Min.ADPE}\t{ele.TotalResult_Max.ADPE}\t{ele.TotalResult_Average.ADPE}\t{ele.TotalResult_SD.ADPE}\t{ele.TotalResult_Min.ADPF}\t{ele.TotalResult_Max.ADPF}\t{ele.TotalResult_Average.ADPF}\t{ele.TotalResult_SD.ADPF}\t{ele.TotalResult_Min.PERT}\t{ele.TotalResult_Max.PERT}\t{ele.TotalResult_Average.PERT}\t{ele.TotalResult_SD.PERT}\t{ele.TotalResult_Min.PENRT}\t{ele.TotalResult_Max.PENRT}\t{ele.TotalResult_Average.PENRT}\t{ele.TotalResult_SD.PENRT}");

                    eleCount++;

                    totalRecord_avg.GWP += ele.TotalResult_Average.GWP;
                    totalRecord_avg.ODP += ele.TotalResult_Average.ODP;
                    totalRecord_avg.AP += ele.TotalResult_Average.AP;
                    totalRecord_avg.EP += ele.TotalResult_Average.EP;
                    totalRecord_avg.POCP += ele.TotalResult_Average.POCP;
                    totalRecord_avg.ADPE += ele.TotalResult_Average.ADPE;
                    totalRecord_avg.ADPF += ele.TotalResult_Average.ADPF;
                    totalRecord_avg.PERT += ele.TotalResult_Average.PERT;
                    totalRecord_avg.PENRT += ele.TotalResult_Average.PENRT;

                    totalRecord_min.GWP += ele.TotalResult_Min.GWP;
                    totalRecord_min.ODP += ele.TotalResult_Min.ODP;
                    totalRecord_min.AP += ele.TotalResult_Min.AP;
                    totalRecord_min.EP += ele.TotalResult_Min.EP;
                    totalRecord_min.POCP += ele.TotalResult_Min.POCP;
                    totalRecord_min.ADPE += ele.TotalResult_Min.ADPE;
                    totalRecord_min.ADPF += ele.TotalResult_Min.ADPF;
                    totalRecord_min.PERT += ele.TotalResult_Min.PERT;
                    totalRecord_min.PENRT += ele.TotalResult_Min.PENRT;

                    totalRecord_max.GWP += ele.TotalResult_Max.GWP;
                    totalRecord_max.ODP += ele.TotalResult_Max.ODP;
                    totalRecord_max.AP += ele.TotalResult_Max.AP;
                    totalRecord_max.EP += ele.TotalResult_Max.EP;
                    totalRecord_max.POCP += ele.TotalResult_Max.POCP;
                    totalRecord_max.ADPE += ele.TotalResult_Max.ADPE;
                    totalRecord_max.ADPF += ele.TotalResult_Max.ADPF;
                    totalRecord_max.PERT += ele.TotalResult_Max.PERT;
                    totalRecord_max.PENRT += ele.TotalResult_Max.PENRT;

                    totalRecord_sd.GWP += Math.Pow(ele.TotalResult_SD.GWP, 2);
                    totalRecord_sd.ODP += Math.Pow(ele.TotalResult_SD.ODP, 2);
                    totalRecord_sd.AP += Math.Pow(ele.TotalResult_SD.AP, 2);
                    totalRecord_sd.EP += Math.Pow(ele.TotalResult_SD.EP, 2);
                    totalRecord_sd.POCP += Math.Pow(ele.TotalResult_SD.POCP, 2);
                    totalRecord_sd.ADPE += Math.Pow(ele.TotalResult_SD.ADPE, 2);
                    totalRecord_sd.ADPF += Math.Pow(ele.TotalResult_SD.ADPF, 2);
                    totalRecord_sd.PERT += Math.Pow(ele.TotalResult_SD.PERT, 2);
                    totalRecord_sd.PENRT += Math.Pow(ele.TotalResult_SD.PENRT, 2);

                }

                totalRecord_sd.GWP = Math.Sqrt(totalRecord_sd.GWP);
                totalRecord_sd.ODP = Math.Sqrt(totalRecord_sd.ODP);
                totalRecord_sd.AP = Math.Sqrt(totalRecord_sd.AP);
                totalRecord_sd.EP = Math.Sqrt(totalRecord_sd.EP);
                totalRecord_sd.POCP = Math.Sqrt(totalRecord_sd.POCP);
                totalRecord_sd.ADPE = Math.Sqrt(totalRecord_sd.ADPE);
                totalRecord_sd.ADPF = Math.Sqrt(totalRecord_sd.ADPF);
                totalRecord_sd.PERT = Math.Sqrt(totalRecord_sd.PERT);
                totalRecord_sd.PENRT = Math.Sqrt(totalRecord_sd.PENRT);

                writer.WriteLine($"Total: \t{totalRecord_min.GWP}\t{totalRecord_max.GWP}\t{totalRecord_avg.GWP}\t{totalRecord_sd.GWP}\t{totalRecord_min.ODP}\t{totalRecord_max.ODP}\t{totalRecord_avg.ODP}\t{totalRecord_sd.ODP}\t{totalRecord_min.AP}\t{totalRecord_max.AP}\t{totalRecord_avg.AP}\t{totalRecord_sd.AP}\t{totalRecord_min.EP}\t{totalRecord_max.EP}\t{totalRecord_avg.EP}\t{totalRecord_sd.EP}\t{totalRecord_min.POCP}\t{totalRecord_max.POCP}\t{totalRecord_avg.POCP}\t{totalRecord_sd.POCP}\t{totalRecord_min.ADPE}\t{totalRecord_max.ADPE}\t{totalRecord_avg.ADPE}\t{totalRecord_sd.ADPE}\t{totalRecord_min.ADPF}\t{totalRecord_max.ADPF}\t{totalRecord_avg.ADPF}\t{totalRecord_sd.ADPF}\t{totalRecord_min.PERT}\t{totalRecord_max.PERT}\t{totalRecord_avg.PERT}\t{totalRecord_sd.PERT}\t{totalRecord_min.PENRT}\t{totalRecord_max.PENRT}\t{totalRecord_avg.PENRT}\t{totalRecord_sd.PENRT}");
            }

            EleLCAResult totalResults = new EleLCAResult();

            totalResults.Name = $"{phase} Total";

            totalResults.Stage = phase;

            totalResults.TotalResult_Min = totalRecord_min;

            totalResults.TotalResult_Max = totalRecord_max;

            totalResults.TotalResult_Average = totalRecord_avg;

            totalResults.TotalResult_SD = totalRecord_sd;

            return totalResults;
        }

        private void PrintTransportRecordResults(List<Dictionary<int, (string, string, List<(double, string, Record)>)>> printResults)
        {
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

            string path = desktopPath + "\\A4Results.txt";

            int eleTotalCount = printResults.Count;

            int eleCount = 1;

            using (var writer = new StreamWriter(path))
            {
                foreach (var ele in printResults)
                {
                    int matTotalCount = ele.Count;

                    int matCount = 1;

                    foreach (var material in ele)
                    {
                        var RecordTotalCount = material.Value.Item3.Count;

                        int RecordCount = 1;

                        foreach (var data in material.Value.Item3)
                        {
                            writer.WriteLine($"Ele\t{eleCount}/{eleTotalCount}\tMaterial{matCount}/{matTotalCount}\t{material.Value.Item1}\t{material.Value.Item2}\tRecord{RecordCount}/{RecordTotalCount}\t{data.Item1}\t{data.Item2}\t{data.Item3.Unit}\t{data.Item3.MaterialType}\t{data.Item3.GWP}");

                            RecordCount++;
                        }
                        matCount++;
                    }

                    eleCount++;
                }
            }
        }
        #endregion

        #region Calculations

        private void CalculateLCA(List<List<MaterialInfo>> eleInfos)
        {
            List<List<MaterialInfo>> InputData = GetQuantities(eleInfos);

            //PrintQuantities(InputData);

            List<EleLCAResult> totalResultList = new List<EleLCAResult>(); //Include A1-A3 Ele total, A4 Ele total, A5 Ele total, All selected stages Ele total

            List<EleLCAResult> A1_A3Results = new List<EleLCAResult>();

            List<EleLCAResult> A4Results = new List<EleLCAResult>();

            List<EleLCAResult> A5Results = new List<EleLCAResult>();

            EleLCAResult A1_A3_EleTotal = new EleLCAResult();

            EleLCAResult A4_EleTotal = new EleLCAResult();

            EleLCAResult A5_EleTotal = new EleLCAResult();

            if (calculate_A1_A3.IsChecked == true)
            {
                A1_A3Results = GetA1_A3Results(InputData);

                A1_A3_EleTotal = WriteResult(A1_A3Results, "A1-A3");

                totalResultList.Add(A1_A3_EleTotal);
            }

            if (calculate_A4.IsChecked == true)
            {
                A4Results = GetA4Results(InputData);

                A4_EleTotal=WriteResult(A4Results, "A4");

                totalResultList.Add(A4_EleTotal);
            }

            if (calculate_A5.IsChecked == true)
            {
                A5Results = GetA5Results(InputData);

                A5_EleTotal=WriteResult(A5Results, "A5");

                totalResultList.Add(A5_EleTotal);
            }

            List<EleLCAResult> AllStageResults = AddTotalResults(A1_A3Results, A4Results, A5Results);

            var AllStage_EleTotal=WriteResult(AllStageResults, AllStageResults[0].Stage);

            string mvd_description = Model.Header.ModelViewDefinition;

            // Regular expression pattern to match the values within [] followed by ViewDefinition
            string pattern = @"ExchangeRequirement\s*\[([^[\]]*)\]";

            // Match the pattern in the line
            Match match = Regex.Match(mvd_description, pattern);

            string ER_Name = "";

            // Check if a match is found
            if (match.Success)
            {
                // Get the value enclosed within square brackets
                ER_Name = match.Groups[1].Value;
            }

            WriteResult(totalResultList, ER_Name);

            totalResultList.Add(AllStage_EleTotal);

            A1_A3Results.Add(A1_A3_EleTotal);

            A4Results.Add(A4_EleTotal);

            A5Results.Add(A5_EleTotal);

            AllStageResults.Add(AllStage_EleTotal);

            /*       EleLCAResult totalResults=new EleLCAResult();

                     var building=Model.Instances.OfType<IfcBuilding>().FirstOrDefault();

                     string buildingLabel = building.EntityLabel.ToString(); 

                     totalResults.IfcLabel = buildingLabel;

                     string buildingName = building.Name.ToString();

                     totalResults.Name = $"Total Results for {buildingName}";

                     totalResults.Description = $"Total Results for {buildingName}";

                     totalResults.Stage = results[0].Stage;

                     totalResults.AverageResults=new Dictionary<string, Record>();

                     totalResults.MinResults = new Dictionary<string, Record>();

                     totalResults.MaxResults = new Dictionary<string, Record>();

                     foreach(var result in results)
                     {
                         totalResults.AverageResults.Add(result.IfcLabel + " - " + result.Name, result.TotalResult_Average);

                         totalResults.MinResults.Add(result.IfcLabel + " - " + result.Name, result.TotalResult_Min);

                         totalResults.MaxResults.Add(result.IfcLabel + " - " + result.Name, result.TotalResult_Max);
                     }
         */
            //GetTotalResults(totalResults);
            //has a problem need to write the min,max,average record value of ele of every single stage and their totals to the file;
            //the same with the display results

            DisplayResults(A1_A3Results, "A1_A3");
            DisplayResults(A4Results, "A4");
            DisplayResults(A5Results, "A5");
            DisplayResults(totalResultList, "Total");
            DisplayResults(AllStageResults, "All Stages");

            MessageBox.Show("The results have been saved to the desktop");

        }

        private void DisplayResults(List<EleLCAResult> results, string v)
        {
            List<Record> Min_Records = new List<Record>();

            List<Record> Max_Records = new List<Record>();

            List<Record> Avg_Records = new List<Record>();

            List<Record> SD_Records = new List<Record>();

           /* Min_Records = results.Select(x => x.TotalResult_Min).ToList();

            Max_Records = results.Select(x => x.TotalResult_Max).ToList();

            Avg_Records = results.Select(x => x.TotalResult_Average).ToList();
*/
            if(results==null || (results.Count==1 && results[0].Name==null))
            {
                return;
            }

            foreach(var ele in results)
            {
                ele.TotalResult_Min.Uid= ele.Name;

                ele.TotalResult_Max.Uid = ele.Name;

                ele.TotalResult_Average.Uid = ele.Name;

                ele.TotalResult_SD.Uid = ele.Name;

                Min_Records.Add(ele.TotalResult_Min);
              
                Max_Records.Add(ele.TotalResult_Max);

                Avg_Records.Add(ele.TotalResult_Average);

                SD_Records.Add(ele.TotalResult_SD);

            }

            switch (v)
            {
                case "A1-A3":
                    A1_A3_Min = Min_Records;
                    A1_A3_Max = Max_Records;
                    A1_A3_Avg = Avg_Records;
                    A1_A3_SD = SD_Records;
                    break;
                case "A4":
                    A4_Min = Min_Records;
                    A4_Max = Max_Records;
                    A4_Avg = Avg_Records;
                    A4_SD = SD_Records;
                    break;
                case "A5":
                    A5_Min = Min_Records;
                    A5_Max = Max_Records;
                    A5_Avg = Avg_Records;
                    A5_SD = SD_Records;
                    break;

                case "Total":
                    Total_Min = Min_Records;
                    Total_Max = Max_Records;
                    Total_Avg = Avg_Records;
                    Total_SD = SD_Records;
                    break;
                case "All Stages":
                    All_Stage_Min = Min_Records;
                    All_Stage_Max = Max_Records;
                    All_Stage_Avg = Avg_Records;
                    All_Stage_SD = SD_Records;
                    break;
            }

        }

        private List<EleLCAResult> AddTotalResults(List<EleLCAResult> A1_A3Results, List<EleLCAResult> A4Results, List<EleLCAResult> A5Results)
        {
            List<EleLCAResult> results = new List<EleLCAResult>();

            bool? a1_a3 = calculate_A1_A3.IsChecked;

            bool? a4 = calculate_A4.IsChecked;

            bool? a5 = calculate_A5.IsChecked;

            foreach (var ele in SelectedElements)
            {
                EleLCAResult totalResult = new EleLCAResult();

                totalResult.Name = ele.Name;

                totalResult.IfcLabel=ele.EntityLabel.ToString();

                totalResult.Description = $"Total Results for ele #{totalResult.IfcLabel} - {totalResult.Name}";

                totalResult.Stage = "";

                totalResult.MaterialResults = new Dictionary<string, List<Record>>();

                totalResult.AverageResults = new Dictionary<string, Record>();

                totalResult.MinResults = new Dictionary<string, Record>();

                totalResult.MaxResults = new Dictionary<string, Record>();

                totalResult.StandardDeviation = new Dictionary<string, Record>();

                if (a1_a3==true)
                {
                    totalResult.Stage+="A1-A3, ";

                    var item = A1_A3Results.Find(x => x.IfcLabel == totalResult.IfcLabel);

                    if (item != null)
                    {                      
                        totalResult.AverageResults.Add("A1-A3",item.TotalResult_Average);

                        totalResult.MinResults.Add("A1-A3", item.TotalResult_Min);
                        
                        totalResult.MaxResults.Add("A1-A3", item.TotalResult_Max);

                        totalResult.StandardDeviation.Add("A1-A3", item.TotalResult_SD);
                    }
                }

                if(a4 == true)
                {
                    totalResult.Stage += "A4, ";

                    var item = A4Results.Find(x => x.IfcLabel == totalResult.IfcLabel);

                    if (item != null)
                    {
                        totalResult.AverageResults.Add("A4", item.TotalResult_Average);

                        totalResult.MinResults.Add("A4", item.TotalResult_Min);

                        totalResult.MaxResults.Add("A4", item.TotalResult_Max);

                        totalResult.StandardDeviation.Add("A4", item.TotalResult_SD);
                    }
                }

                if (a5 == true)
                {
                    totalResult.Stage += "A5";

                    var item = A5Results.Find(x => x.IfcLabel == totalResult.IfcLabel);

                    if (item != null)
                    {
                        totalResult.AverageResults.Add("A5", item.TotalResult_Average);

                        totalResult.MinResults.Add("A5", item.TotalResult_Min);

                        totalResult.MaxResults.Add("A5", item.TotalResult_Max);

                        totalResult.StandardDeviation.Add("A5", item.TotalResult_SD);
                    }
                }

                GetTotalResults(totalResult);

                results.Add(totalResult);
                
            }

            if (calculate_A5.IsChecked == true && A5Results.Count == 1)
            {
                results.Add(A5Results[0]);
            }

            return results;         
        }

        #region A1-A3

        private List<EleLCAResult> GetA1_A3Results(List<List<MaterialInfo>> eleInfos)
        {
            List<EleLCAResult> results = new List<EleLCAResult>();

            foreach (var ele in eleInfos)
            {
                EleLCAResult eleResult = new EleLCAResult();

                eleResult.MaterialResults = new Dictionary<string, List<Record>>();

                eleResult.Name = ele[0].Name;

                eleResult.IfcLabel = ele[0].IfcLabel;

                eleResult.Stage = "A1-A3";

                eleResult.Description = $"Total of {ele.Count} materials:; ";

                eleResult.materialCount = ele.Count;

                int matCount = 1;

                foreach (var material in ele)
                {
                    List<Record> records = new List<Record>();

                    if (material.MaterialRecords.Count == 0 || material.unit == "N/A")
                    {
                        eleResult.Description += $"No LCA/Quantity data found for material {matCount};";

                        continue;
                    }

                    foreach (Record record in material.MaterialRecords)
                    {
                        Record recordResult = SingleMaterialRecordResult(record, material.quantity, material.unit);

                        records.Add(recordResult);
                    }

                    eleResult.MaterialResults.Add(matCount.ToString()+"-"+material.ID, records);

                    eleResult.Description += $"Total of {material.MaterialRecords.Count} sampled for material {matCount};";

                    matCount++;
                }

                CalculateStaticticResults(eleResult);

                results.Add(eleResult);
            }

            //WriteLCAResult(results,"A1-A3");

            return results;
        }

        private Record SingleMaterialRecordResult(Record record, double quantity, string unit)
        {
            UnitConverter unitConverter = new UnitConverter(unit, record.Unit);

            double convertedQuantity = unitConverter.Convert(quantity);

            Record recordResult = new Record();

            recordResult.Uid = record.Uid;

            recordResult.Unit = record.Unit;

            recordResult.MaterialType = record.MaterialType;

            recordResult.GWP = record.GWP * convertedQuantity;

            recordResult.ODP = record.ODP * convertedQuantity;

            recordResult.AP = record.AP * convertedQuantity;

            recordResult.EP = record.EP * convertedQuantity;

            recordResult.POCP = record.POCP * convertedQuantity;

            recordResult.ADPE = record.ADPE * convertedQuantity;

            recordResult.ADPF = record.ADPF * convertedQuantity;

            recordResult.PERT = record.PERT * convertedQuantity;

            recordResult.PENRT = record.PENRT * convertedQuantity;

            recordResult.Address = record.Address;

            return recordResult;
        }

        #region RetriveRecordData

        private List<MaterialInfo> RetrieveLCAInfo(string ele, SQLDBConnect connect, int samplingNum, int samplingMode)
        {
            string[] infos = ele.Split('\t');

            string UID_string = infos[4];

            string[] UIDs = UID_string.Split(';');

            List<MaterialInfo> materialInfos = new List<MaterialInfo>();

            foreach (var uid in UIDs)
            {
                List<Record> records = new List<Record>();

                string[] specificInfos = uid.Split('-');

                string[] truncatedArray = new string[specificInfos.Length - 2];

                Array.Copy(specificInfos, truncatedArray, specificInfos.Length - 2);

                string realID = string.Join("-", truncatedArray);

                if (specificInfos.Length == 7 && Guid.TryParse(realID, out Guid id))
                {
                    if (CheckIfAverageData(id, connect))
                    {
                        List<Record> averageDataRecords = RetriveSamplingData(id, connect, samplingNum, samplingMode);

                        if (averageDataRecords != null)
                        {
                            records.AddRange(averageDataRecords);
                        }
                    }
                    else
                    {
                        Record record = RetriveSpecificData(id, connect);

                        if (record != null)
                        {
                            records.Add(record);
                        }
                    }
                }
                else
                {
                    List<Record> averageDataRecords = SearchForAverageData(uid, connect, samplingNum, samplingMode);

                    if (averageDataRecords != null)
                    {
                        records.AddRange(averageDataRecords);
                    }
                }

                MaterialInfo newElement = new MaterialInfo() { IfcLabel = infos[0], Name = infos[1], IfcType = infos[2], ElementType = infos[3], ID = uid, MaterialRecords = records };

                materialInfos.Add(newElement);
            }

            return materialInfos;
        }

        private bool CheckIfAverageData(Guid id, SQLDBConnect connect)
        {
            string sql = $"SELECT COUNT(*) FROM [dbo].[AverageMaterialTable] WHERE Material_ID = '{id.ToString()}'";

            SqlCommand query = connect.Query(sql);

            int count = (int)query.ExecuteScalar();

            if (count > 0)
            {
                return true;
            }

            return false;
        }

        private List<Record> SearchForAverageData(string uid, SQLDBConnect connect, int samplingNum, int samplingMode)
        {
            //Category, Manufacturer, Model, Year, RegionLevel, Region, Thickness
            string[] details = uid.Split('-');

            //There is a problem regarding the material Types being assigned

            string sql = $"SELECT [dbo].[MaterialTable].[Material_ID], [dbo].[MaterialTable].[Manufacturer_ID], [dbo].[FactoryList].[Factory_CityName] From [dbo].[MaterialTable] " +
                $"INNER JOIN [dbo].[FactoryList] ON [dbo].[FactoryList].[Factory_ID]=[dbo].[MaterialTable].[Factory_ID]" +
                $"INNER JOIN [dbo].[LCA_GeneralInfo] ON [dbo].[LCA_GeneralInfo].[Material_ID]=[dbo].[MaterialTable].[Material_ID]" +
                $"INNER JOIN [dbo].[StateCodeList] ON [dbo].[StateCodeList].[StateCode]=[dbo].[FactoryList].[Factory_StateCode]" +
                $"INNER JOIN [dbo].[CountryCodeList] ON [dbo].[CountryCodeList].[CountryCode]=[dbo].[FactoryList].[Factory_CountryCode]" +
                $"WHERE [dbo].[MaterialTable].[Material_Type] LIKE '%{GetSynonym(details[0])}%'";

            if (details[1] != "XXX")
            {
                sql += $"AND [dbo].[MaterialTable].[Manufacturer_Name]='{details[1]}'";
            }

            if (details[2] != "XXX")
            {
                sql += $"AND ([dbo].[MaterialTable].[Model_Name]='{details[2]}' OR [dbo].[MaterialTable].[Model_Code]='{details[2]}')";
            }

            if (details[3] != "XXX")
            { 
                sql += $"AND ([dbo].[LCA_GeneralInfo].[ValidFrom] <= '{details[3]}-01-01' AND [dbo].[LCA_GeneralInfo].[ValidUntil] >= '{details[3]}-12-31')";

            }

            if (details[4] != "XXX" && details[5] != "XXX")
            {
                if (details[4] == "Country" || details[4] == "Nation")
                {
                    sql += $"AND ([dbo].[CountryCodeList].[CountryName]='{details[5]}' OR [dbo].[CountryCodeList].[CountryCode]='{details[5]}')";
                }
                else if (details[4] == "State" || details[4] == "Province")
                {
                    sql += $"AND ([dbo].[StateCodeList].[StateName]='{details[5]}' OR [dbo].[StateCodeList].[StateCode]='{details[5]}')";
                }
                else if (details[4] == "City" || details[4] == "Town")
                {
                    sql += $"AND [dbo].[FactoryList].[Factory_CityName]='{details[5]}'";
                }
            }

            sql += $";";

            SqlCommand query = connect.Query(sql);

            List<(Guid, int, string)> records = new List<(Guid, int, string)>();

            using (SqlDataReader reader = query.ExecuteReader())
            {
                while (reader.Read())
                {
                    records.Add((reader.GetGuid(0), reader.GetInt32(1), reader.GetString(2)));
                }
            }

            if (records.Count > 0)
            {
                (bool, Guid) result = connect.InsertData(records, details[0], details[3]);

                if (result.Item1)
                {
                    return RetriveSamplingData(result.Item2, connect, samplingNum, samplingMode);
                }
            }

            return null;
        }

        private string GetSynonym(string type)
        {
            string caseInsensitive = type.ToLower().Replace(" ","");

            switch (caseInsensitive)
            {
                case "window":
                case "windows":
                    return "Window";

                case "door":
                case "doors":
                    return "Door";

                case "timber":
                case "wood":
                case "woods":
                case "timbers":
                    return "Timber";

                case "concrete":
                case "concretes":
                case "readymixconcrete":
                case "readymixconcretes":
                    return "ReadyMixConcrete";

                case "steel":
                case "steels":
                    return "Steel";

                case "metal":
                case "metals":
                    return "Steel General";

                case "reinforcement":
                case "reinforcements":
                    return "Steel Reinforcement";

                case "furniture":
                case "furnitures":
                    return "Furniture";

                case "brick":
                case "bricks":
                case "masonry":
                    return "Masonry";

                case "glass":
                case "glasses":
                    return "Glass";

                case "plaster":
                case "plasters":
                case "plasterboard":
                case "plasterboards":
                case "gypsumboard":
                case "gypsumboards":
                case "drywall":
                case "drywalls":
                    return "Plaster board";

                case "mdf":
                case "mdfs":
                case "mediumdensityfiberboard":
                case "mediumdensityfiberboards":
                    return "MDF";

                case "particleboard":
                case "particleboards":
                    return "Particle Board";

                case "insulation":
                case "insulations":
                case "insulationpanel":
                case "insulationpanels":
                    return "Insulation Panels";

                case "cladding":
                case "claddings":
                case "claddingsystem":
                case "claddingsystems":
                    return "Cladding";

                case "celling":
                case "cellings":
                case "cellingsystem":
                case "cellingsystems":
                    return "Celling System";

                case "floor":
                case "floors":
                case "flooring":
                case "floorings":
                case "flooringsystem":
                case "flooringsystems":
                    return "Floor System";

                case "ceramic":
                case "ceramics":
                    return "Ceramic";

                case "Aluminium":
                case "aluminiums":
                    return "Aluminium";

                default:
                    return type;

            }
        }

        private List<Record> RetriveSamplingData(Guid id, SQLDBConnect connect, int samplingNum, int samplingMode)
        {
            List<Guid> Guids = GetSampleGuids(id, connect, samplingNum, samplingMode);

            List<Record> records = new List<Record>();

            if (Guids == null)
            {
                return null;
            }

            foreach (var Guid in Guids)
            {
                records.Add(RetriveSpecificData(Guid, connect));
            }

            return records;
        }

        private List<Guid> GetSampleGuids(Guid id, SQLDBConnect connect, int samplingNum, int samplingMode)
        {
            string sql = $"SELECT [dbo].[AverageDataTable].[Material_ID], Unit FROM [dbo].[AverageDataTable] " +
                $"INNER JOIN [dbo].[LCA_GeneralInfo] ON [dbo].[LCA_GeneralInfo].[Material_ID]=[dbo].[AverageDataTable].[Material_ID]" +
                $"WHERE AverageMaterial_ID = '{id.ToString()}';";

            SqlCommand query = connect.Query(sql);

            List<(Guid, string)> ids = new List<(Guid, string)>();

            using (SqlDataReader reader = query.ExecuteReader())
            {
                while (reader.Read())
                {
                    ids.Add((reader.GetGuid(0), reader.GetString(1)));
                }
            }

            if (ids.Count > 0)
            {
                return SamplingGuids(ids, samplingNum, samplingMode);
            }

            return null;
        }

        private List<Guid> SamplingGuids(List<(Guid, string)> data, int samplingNum, int samplingMode)
        {
            Dictionary<string, List<Guid>> IdsByUnit = new Dictionary<string, List<Guid>>();

            foreach (var pair in data)
            {
                if (IdsByUnit.ContainsKey(pair.Item2))
                {
                    IdsByUnit[pair.Item2].Add(pair.Item1);
                }
                else
                {
                    List<Guid> newGuids = new List<Guid>();

                    newGuids.Add(pair.Item1);

                    IdsByUnit.Add(pair.Item2, newGuids);
                }
            }

            Dictionary<string, List<Guid>> IdsByUnitCategory = new Dictionary<string, List<Guid>>();

            foreach (var pair in IdsByUnit)
            {
                string category = DetermineUnitCategory(pair.Key);

                if (IdsByUnitCategory.ContainsKey(category))
                {
                    IdsByUnitCategory[category].AddRange(pair.Value);
                }
                else
                {
                    IdsByUnitCategory.Add(category, pair.Value);
                }

            }

            List<Guid> ids = IdsByUnit.OrderByDescending(pair => pair.Value.Count).FirstOrDefault().Value;

            int count = ids.Count;

            if (samplingMode == 1)
            {

                Random random = new Random();

                List<Guid> sampleGuids = new List<Guid>();

                if (count <= samplingNum)
                {
                    return ids;
                }

                while (sampleGuids.Count < samplingNum)
                {
                    int index = random.Next(0, count);

                    if (!sampleGuids.Contains(ids[index]))
                    {
                        sampleGuids.Add(ids[index]);
                    }
                }
                return sampleGuids;
            }
            else if (samplingMode == 2)
            {
                int sampleSize = (int)Math.Ceiling(count * samplingNum / 100.0);

                if (count <= sampleSize)
                {
                    return ids;
                }

                Random random = new Random();

                List<Guid> sampleGuids = new List<Guid>();

                while (sampleGuids.Count < sampleSize)
                {
                    int index = random.Next(0, count);

                    if (!sampleGuids.Contains(ids[index]))
                    {
                        sampleGuids.Add(ids[index]);
                    }
                }
                return sampleGuids;
            }

            return null;
        }

        private Record RetriveSpecificData(Guid id, SQLDBConnect connect)
        {
            string sql = $"SELECT Unit,GWP,ODP,AP,EP,POCP,ADPE,ADPF,PERT,PENRT,Factory_StreetAddress,Factory_Name,Factory_CityName,Factory_StateCode,Factory_CountryCode,Material_Type FROM [dbo].[LCA_Datasets] " +
                $"INNER JOIN [dbo].[LCA_GeneralInfo] ON [dbo].[LCA_DataSets].[Material_ID]=[dbo].[LCA_GeneralInfo].[Material_ID] " +
                $"INNER JOIN [dbo].[MaterialTable] ON [dbo].[LCA_DataSets].[Material_ID]=[dbo].[MaterialTable].[Material_ID]" +
                $"INNER JOIN [dbo].[FactoryList] ON [dbo].[MaterialTable].[Factory_ID]=[dbo].[FactoryList].[Factory_ID]" +
                $"WHERE [dbo].[LCA_DataSets].[Material_ID] = '{id.ToString()}';";

            SqlCommand query = connect.Query(sql);

            using (SqlDataReader reader = query.ExecuteReader())
            {
                if (reader.Read())
                {
                    Record record = new Record()
                    {
                        Uid = id.ToString(),
                        Unit = reader.GetString(0).Trim(),
                        GWP = reader.GetDouble(1),
                        ODP = reader.GetDouble(2),
                        AP = reader.GetDouble(3),
                        EP = reader.GetDouble(4),
                        POCP = reader.GetDouble(5),
                        ADPE = reader.GetDouble(6),
                        ADPF = reader.GetDouble(7),
                        PERT = reader.GetDouble(8),
                        PENRT = reader.GetDouble(9),
                        Address = GetAddress(reader.GetString(11).Trim(), reader.GetString(10).Trim(), reader.GetString(12).Trim(), reader.GetString(13).Trim(), reader.GetString(14).Trim()),
                        MaterialType = reader.GetString(15)
                    };

                    return record;
                }
                else
                {
                    MessageBox.Show("Error!, No record can be found!");

                    return null;
                }

            }

        }

        private string GetAddress(string FactoryName, string StreetAddress, string CityName, string StateCode, string CountryCode)
        {
            if (string.IsNullOrEmpty(StreetAddress))
            {
/*                if (StateCode == "N/A")
                {
                    return string.Join(", ", FactoryName, CityName, CountryCode);
                }*/

                return string.Join(",", FactoryName, CityName, StateCode, CountryCode);
            }
/*
            if (StateCode == "N/A")
            {
                return string.Join(", ", StreetAddress, CityName, CountryCode);
            }*/

            return string.Join(",", StreetAddress, CityName, StateCode, CountryCode);
        }

        #endregion

        #endregion

        #region GetStatisticResults

        private void CalculateStaticticResults(EleLCAResult eleResult)
        {
            eleResult.MinResults = new Dictionary<string, Record>();

            eleResult.MaxResults = new Dictionary<string, Record>();

            eleResult.AverageResults = new Dictionary<string, Record>();

            eleResult.StandardDeviation = new Dictionary<string, Record>();

            foreach (var material in eleResult.MaterialResults)
            {
                Record minRecord = GetMin(material.Value);

                Record maxRecord = GetMax(material.Value);

                Record avergRecord = GetAverage(material.Value);

                Record sdRecord = GetStandardDeviation(material.Value, avergRecord);

                eleResult.MinResults.Add(material.Key, minRecord);

                eleResult.MaxResults.Add(material.Key, maxRecord);

                eleResult.AverageResults.Add(material.Key, avergRecord);

                eleResult.StandardDeviation.Add(material.Key, sdRecord);

            }

            GetTotalResults(eleResult);
        }

        private Record GetStandardDeviation(List<Record> value, Record averageRecord)
        {
            Record sdRecord = new Record()
            {
                Uid = value[0].Uid,
                Unit = value[0].Unit,
                Address = value[0].Address,
                GWP = CalculateSD(value.Select(x => x.GWP).ToList(), averageRecord.GWP),
                ODP = CalculateSD(value.Select(x => x.ODP).ToList(), averageRecord.ODP),
                AP = CalculateSD(value.Select(x => x.AP).ToList(), averageRecord.AP),
                EP = CalculateSD(value.Select(x => x.EP).ToList(), averageRecord.EP),
                POCP = CalculateSD(value.Select(x => x.POCP).ToList(), averageRecord.POCP),
                ADPE = CalculateSD(value.Select(x => x.ADPE).ToList(), averageRecord.ADPE),
                ADPF = CalculateSD(value.Select(x => x.ADPF).ToList(), averageRecord.ADPF),
                PERT = CalculateSD(value.Select(x => x.PERT).ToList(), averageRecord.PERT),
                PENRT = CalculateSD(value.Select(x => x.PENRT).ToList(), averageRecord.PENRT),
            };

            return sdRecord;
        }

        private double CalculateSD(List<double> list, double averageValue)
        {
            double variance = list.Sum(x => Math.Pow(x - averageValue, 2)) / list.Count;

            return Math.Sqrt(variance);
        }

        private Record GetAverage(List<Record> value)
        {
            var averageRecord = new Record()
            {
                Uid = value[0].Uid,
                Unit = value[0].Unit,
                Address = value[0].Address,
                MaterialType = value[0].MaterialType,
                GWP = value.Average(x => x.GWP),
                ODP = value.Average(x => x.ODP),
                AP = value.Average(x => x.AP),
                EP = value.Average(x => x.EP),
                POCP = value.Average(x => x.POCP),
                ADPE = value.Average(x => x.ADPE),
                ADPF = value.Average(x => x.ADPF),
                PERT = value.Average(x => x.PERT),
                PENRT = value.Average(x => x.PENRT),
            };

            return averageRecord;
        }

        private Record GetMax(List<Record> value)
        {
            Record maxRecord = new Record()
            {
                Uid = value[0].Uid,
                Unit = value[0].Unit,
                Address = value[0].Address,
                MaterialType = value[0].MaterialType,
                GWP = value.Max(x => x.GWP),
                ODP = value.Max(x => x.ODP),
                AP = value.Max(x => x.AP),
                EP = value.Max(x => x.EP),
                POCP = value.Max(x => x.POCP),
                ADPE = value.Max(x => x.ADPE),
                ADPF = value.Max(x => x.ADPF),
                PERT = value.Max(x => x.PERT),
                PENRT = value.Max(x => x.PENRT),
            };
            return maxRecord;
        }

        private Record GetMin(List<Record> value)
        {
            Record minRecord = new Record()
            {
                Uid = value[0].Uid,
                Unit = value[0].Unit,
                Address = value[0].Address,
                MaterialType = value[0].MaterialType,
                GWP = value.Min(x => x.GWP),
                ODP = value.Min(x => x.ODP),
                AP = value.Min(x => x.AP),
                EP = value.Min(x => x.EP),
                POCP = value.Min(x => x.POCP),
                ADPE = value.Min(x => x.ADPE),
                ADPF = value.Min(x => x.ADPF),
                PERT = value.Min(x => x.PERT),
                PENRT = value.Min(x => x.PENRT),
            };

            return minRecord;
        }

        private void GetTotalResults(EleLCAResult eleResult)
        {
            eleResult.TotalResult_Average = new Record()
            {
                Uid = "Total",
                Unit = "N/A",
                Address = "N/A",
                MaterialType = "N/A",
                GWP = eleResult.AverageResults.Sum(x => x.Value.GWP),
                ODP = eleResult.AverageResults.Sum(x => x.Value.ODP),
                AP = eleResult.AverageResults.Sum(x => x.Value.AP),
                EP = eleResult.AverageResults.Sum(x => x.Value.EP),
                POCP = eleResult.AverageResults.Sum(x => x.Value.POCP),
                ADPE = eleResult.AverageResults.Sum(x => x.Value.ADPE),
                ADPF = eleResult.AverageResults.Sum(x => x.Value.ADPF),
                PERT = eleResult.AverageResults.Sum(x => x.Value.PERT),
                PENRT = eleResult.AverageResults.Sum(x => x.Value.PENRT)
            };

            eleResult.TotalResult_Max = new Record()
            {
                Uid = "Max",
                Unit = "N/A",
                Address = "N/A",
                MaterialType = "N/A",
                GWP = eleResult.MaxResults.Sum(x => x.Value.GWP),
                ODP = eleResult.MaxResults.Sum(x => x.Value.ODP),
                AP = eleResult.MaxResults.Sum(x => x.Value.AP),
                EP = eleResult.MaxResults.Sum(x => x.Value.EP),
                POCP = eleResult.MaxResults.Sum(x => x.Value.POCP),
                ADPE = eleResult.MaxResults.Sum(x => x.Value.ADPE),
                ADPF = eleResult.MaxResults.Sum(x => x.Value.ADPF),
                PERT = eleResult.MaxResults.Sum(x => x.Value.PERT),
                PENRT = eleResult.MaxResults.Sum(x => x.Value.PENRT),
            };

            eleResult.TotalResult_Min = new Record()
            {
                Uid = "Min",
                Unit = "N/A",
                Address = "N/A",
                MaterialType = "N/A",
                GWP = eleResult.MinResults.Sum(x => x.Value.GWP),
                ODP = eleResult.MinResults.Sum(x => x.Value.ODP),
                AP = eleResult.MinResults.Sum(x => x.Value.AP),
                EP = eleResult.MinResults.Sum(x => x.Value.EP),
                POCP = eleResult.MinResults.Sum(x => x.Value.POCP),
                ADPE = eleResult.MinResults.Sum(x => x.Value.ADPE),
                ADPF = eleResult.MinResults.Sum(x => x.Value.ADPF),
                PERT = eleResult.MinResults.Sum(x => x.Value.PERT),
                PENRT = eleResult.MinResults.Sum(x => x.Value.PENRT),
            };

            List<Record> sdRecords = eleResult.StandardDeviation.Values.ToList();

            eleResult.TotalResult_SD = new Record()
            {
                Address = "N/A",
                MaterialType = "N/A",
                Unit = "N/A",
                Uid = "SD",
                GWP = Math.Sqrt(sdRecords.Sum(x => Math.Pow(x.GWP, 2))),
                ODP = Math.Sqrt(sdRecords.Sum(x => Math.Pow(x.ODP, 2))),
                AP = Math.Sqrt(sdRecords.Sum(x => Math.Pow(x.AP, 2))),
                EP = Math.Sqrt(sdRecords.Sum(x => Math.Pow(x.EP, 2))),
                POCP = Math.Sqrt(sdRecords.Sum(x => Math.Pow(x.POCP, 2))),
                ADPE = Math.Sqrt(sdRecords.Sum(x => Math.Pow(x.ADPE, 2))),
                ADPF = Math.Sqrt(sdRecords.Sum(x => Math.Pow(x.ADPF, 2))),
                PERT = Math.Sqrt(sdRecords.Sum(x => Math.Pow(x.PERT, 2))),
                PENRT = Math.Sqrt(sdRecords.Sum(x => Math.Pow(x.PENRT, 2))),
            };
        }

        private double aT(int s)
        {
            Random random = new Random(s);

            int i = random.Next(1, 100);

            int j = random.Next(1, 500);

            int c = random.Next(1, 4);

            Random random_Percentage = new Random();

            int x = random_Percentage.Next(i, i + j);

            int y = random_Percentage.Next(50, c * 50);

            int z= random_Percentage.Next(150, c * 150);

            double d = ((double)x / 35000) * ((double)y / z);

            return d;
        }

        #endregion

        #region A4

        private List<EleLCAResult> GetA4Results(List<List<MaterialInfo>> eleInfos)
        {
            List<EleLCAResult> results = new List<EleLCAResult>();

            string projectAddress = GetProjectAddress();

            if (projectAddress == null)
            {
                MessageBox.Show("No project address found");

                return null;
            }

            string[] projectAddressString = projectAddress.Split(',');

            int eleCount = 1;

            List<Dictionary<int, (string, string, List<(double, string, Record)>)>> printResults = new List<Dictionary<int, (string, string, List<(double, string, Record)>)>> ();

            foreach (var ele in eleInfos)
            {
                EleLCAResult eleResult = new EleLCAResult();

                eleResult.MaterialResults = new Dictionary<string, List<Record>>();

                eleResult.Name = ele[0].Name;

                eleResult.IfcLabel = ele[0].IfcLabel;

                eleResult.Stage = "A4";

                eleResult.Description = $"Total of {ele.Count} materials:; ";

                eleResult.materialCount = ele.Count;

                int matCount = 1;

                Dictionary<int,(string,string,List<(double,string,Record)>)> keyValuePairs = new Dictionary<int, (string, string, List<(double, string, Record)>)>();

                foreach (var material in ele)
                {
                    if (material.MaterialRecords.Count == 0 || material.unit == "N/A")
                    {
                        eleResult.Description += $"No LCA/Quantity/Distance data found for material {matCount}; ";

                        continue;
                    }

                    List<Record> records = new List<Record>();

                    List<(double, string, Record)> printData = new List<(double, string, Record)>();

                    string description = "";

                    foreach (var record in material.MaterialRecords)
                    {                       
                        string a = null;

                        a = record.Address;

                        string[] address = a.Split(',');

                        string[] GeoScope = { address[address.Length - 2], address.Last()};

                        if (address.Last().Trim().ToLower() == projectAddressString.Last().Trim().ToLower() || GetCountryName(address.Last().Trim()).ToLower() == projectAddressString.Last().Trim().ToLower())
                        {
                            (double, string, Record) distanceData = GetDistanceDataDomestic(projectAddress, a, record.MaterialType, GeoScope);

                            printData.Add(distanceData);

                            Record TransportResult = CalculateA4(distanceData, material);

                            if(TransportResult == null)
                            {
                                description += $"No quantity found for {record.Uid}; ";

                                continue;
                            }

                            records.Add(TransportResult);
                            
                            description += $"Result found for {record.Uid}, Transport {material.quantity} {material.unit} for {distanceData.Item1} {distanceData.Item2} by {distanceData.Item3.MaterialType}; ";

                        }
                        else
                        {
                            List<(double, string, Record)> distanceData = GetDistanceDataInternational(projectAddress, a, record.MaterialType);

                            printData.AddRange(distanceData);

                            if (distanceData == null)
                            {
                                description += $"No result found for {record.Uid}; ";

                                continue;
                            }

                            List<Record> TransportResults = new List<Record>();

                            description += $"Result found for {record.Uid}, Transport {material.quantity} {material.unit} ";

                            foreach (var data in distanceData)
                            {
                                Record TransportResult = CalculateA4(data, material);

                                if (TransportResult == null)
                                {
                                    description += $"No quantity found for {record.Uid}; ";

                                    continue;
                                }

                                TransportResults.Add(TransportResult);

                                description += $"{data.Item1} {data.Item2} by {data.Item3.MaterialType}, ";
                            }

                            description=description.Remove(description.Length - 2);

                            description += "; ";

                            Record SumResult = new Record()
                            {
                                Uid=record.Uid,

                                MaterialType = record.MaterialType,

                                Unit = record.Unit,

                                Address = record.Address,

                                GWP = TransportResults.Sum(x => x.GWP),

                                ODP = TransportResults.Sum(x => x.ODP),

                                AP = TransportResults.Sum(x => x.AP),

                                EP = TransportResults.Sum(x => x.EP),

                                POCP = TransportResults.Sum(x => x.POCP),

                                ADPE = TransportResults.Sum(x => x.ADPE),

                                ADPF = TransportResults.Sum(x => x.ADPF),

                                PERT = TransportResults.Sum(x => x.PERT),

                                PENRT = TransportResults.Sum(x => x.PENRT),
                            };

                            records.Add(SumResult);
                        }                    
                    }

                    //description += "; ";

                    eleResult.MaterialResults.Add(matCount.ToString()+"-"+material.ID, records);

                    eleResult.Description += description;

                    keyValuePairs.Add(matCount,(material.Name,material.ID,printData));

                    matCount++;
                }

                printResults.Add(keyValuePairs);

                eleCount++;
               
                CalculateStaticticResults(eleResult);

                results.Add(eleResult);
            }

            //PrintTransportDataResults(printResults);

            //WriteLCAResult(results, "A4");

            return results;
        }

        private string GetCountryName(string countryCode)
        {
            string sql= $"SELECT CountryName FROM [dbo].[CountryCodeList] WHERE CountryCode='{countryCode}';";

            SqlCommand query = connect.Query(sql);

            string countryName = "";

            using (SqlDataReader reader = query.ExecuteReader())
            {
                if (reader.Read())
                {
                    countryName = reader.GetString(0);
                }
            }

            return countryName;
        }

        private Record CalculateA4((double, string, Record) A4data, MaterialInfo materialInfo)
        {
            string[] Units = A4data.Item3.Unit.Split('*');

            string targetDistanceUnit = null;

            string targetQuantityUnit = null;

            if (Units.Length == 2)
            {
                if (DetermineUnitCategory(Units[1]) == "Length")
                {
                    targetDistanceUnit = Units[1];

                    targetQuantityUnit = Units[0];
                }
                else
                {
                    targetDistanceUnit = Units[0];

                    targetQuantityUnit = Units[1];
                }

            }
            else
            {
                targetDistanceUnit = Units[0];

                targetQuantityUnit = "N/A";
            }

            double quantity = materialInfo.quantity;

            string originQuantityUnit = materialInfo.unit;

            string originQuantityUnitCategory = DetermineUnitCategory(materialInfo.unit);

            string targetQuantityUnitCategory = DetermineUnitCategory(targetQuantityUnit);

            if (targetQuantityUnit != "N/A")
            {
                if (originQuantityUnitCategory != targetQuantityUnitCategory)
                {
                    if (targetQuantityUnitCategory == "Weight" && (materialInfo.IfcType == "IfcWindow" || materialInfo.IfcType == "IfcDoor" || materialInfo.IfcType == "IfcFurniture"))
                    {
                        if (materialInfo.IfcType == "IfcWindow")
                        {
                            List<double> weights = new List<double>();

                            foreach (var record in materialInfo.MaterialRecords)
                            {
                                double weight = RetriveWeight(record.Uid, "Window");

                                UnitConverter unitConverter_W = new UnitConverter(materialInfo.unit, record.Unit);

                                weight = unitConverter_W.Convert(weight);

                                weights.Add(weight);
                            }

                            quantity = weights.Average() * quantity;

                            originQuantityUnit = "kg";

                        }
                        else if (materialInfo.IfcType == "IfcDoor")
                        {
                            List<double> weights = new List<double>();

                            foreach (var record in materialInfo.MaterialRecords)
                            {

                                double weight = RetriveWeight(record.Uid, "Door");

                                UnitConverter unitConverter_W = new UnitConverter(materialInfo.unit, record.Unit);

                                weight = unitConverter_W.Convert(weight);

                                weights.Add(weight);
                            }

                            quantity = weights.Average() * quantity;

                            originQuantityUnit = "kg";
                        }
                        else if (materialInfo.IfcType == "IfcFurniture")
                        {
                            List<double> weights = new List<double>();

                            foreach (var record in materialInfo.MaterialRecords)
                            {
                                double weight = RetriveWeight(record.Uid, "Funiture");

                                UnitConverter unitConverter_W = new UnitConverter(materialInfo.unit, record.Unit);

                                weight = unitConverter_W.Convert(weight);

                                weights.Add(weight);
                            }

                            quantity = weights.Average() * quantity;

                            originQuantityUnit = "kg";
                        }                     
                    }
                    else if (materialInfo.MaterialRecords[0].MaterialType == "Glass" || materialInfo.MaterialRecords[0].MaterialType == "Ceramic Tile" || materialInfo.MaterialRecords[0].MaterialType == "Plaster board")
                    {
                        List<double> weights = new List<double>();

                        foreach (var record in materialInfo.MaterialRecords)
                        {
                            double weight = RetriveWeight(record.Uid, materialInfo.MaterialRecords[0].MaterialType);

                            UnitConverter unitConverter_W = new UnitConverter(materialInfo.unit, record.Unit);

                            weight = unitConverter_W.Convert(weight);

                            weights.Add(weight);
                        }

                        quantity = weights.Average() * quantity;

                        originQuantityUnit = "kg";
                    }
                    else
                    {
                        string[] stringArray = materialInfo.ID.Split('-');

                        string densityString = stringArray.Last();

                        string thicknessString = stringArray[stringArray.Length - 2];

                        (double, string) quantityData = (0, "N/A");

                        double density = 0;

                        double thickness = 0;

                        bool t_S = double.TryParse(thicknessString, out thickness);

                        bool d_S = double.TryParse(densityString, out density);

                        quantityData = GetQuantityByUnit(targetQuantityUnitCategory, materialInfo.IfcLabel, materialInfo.IfcType, thickness, density);

                        if (quantityData.Item2 == "N/A")
                        {
                            MessageBox.Show("Error! No quantity data found"); //IfcMemeber problem

                            //Project Address Problem// solved?

                            return null;
                        }

                        quantity = quantityData.Item1;

                        originQuantityUnit = quantityData.Item2;

                    }
                }

                UnitConverter unitConverter_Q = new UnitConverter(originQuantityUnit, targetQuantityUnit);

                quantity = unitConverter_Q.Convert(quantity);
            }
            else
            {
                quantity = 1;
            }

            UnitConverter unitConverter_D = new UnitConverter(A4data.Item2, targetDistanceUnit);

            double distance = unitConverter_D.Convert(A4data.Item1);

            Record newRecord = new Record();
            {
                newRecord.GWP = A4data.Item3.GWP * distance * quantity;

                newRecord.ODP = A4data.Item3.ODP * distance * quantity;

                newRecord.AP = A4data.Item3.AP * distance * quantity;

                newRecord.EP = A4data.Item3.EP * distance * quantity;

                newRecord.POCP = A4data.Item3.POCP * distance * quantity;

                newRecord.ADPE = A4data.Item3.ADPE * distance * quantity;

                newRecord.ADPF = A4data.Item3.ADPF * distance * quantity;

                newRecord.PERT = A4data.Item3.PERT * distance * quantity;

                newRecord.PENRT = A4data.Item3.PENRT * distance * quantity;
            }
            
            return newRecord;
        }

        #region GetAddress

        private string[] GetPortAddress(string supplierAddress)
        {
            string[] address = supplierAddress.Split(',');

            string sql = $"SELECT Port_StreetAddress,Port_CityName,Port_StateCode,Port_CountryCode,Port_Code FROM [dbo].[PortAddress]" +
                $"INNER JOIN [dbo].[CountryCodeList] ON [dbo].[PortAddress].[Port_CountryCode]=[dbo].[CountryCodeList].[CountryCode]";

            if (address[address.Length - 2].Trim() == "N/A")
            {
                sql += $" WHERE Port_CityName = '{address[address.Length - 2].Trim()}' AND (Port_CountryCode='{address.Last().Trim()}' OR CountryName='{address.Last().Trim()}');";
            }
            else
            {
                sql += $" WHERE Port_StateCode = '{address[address.Length - 2].Trim()}' AND (Port_CountryCode='{address.Last().Trim()}' OR CountryName='{address.Last().Trim()}');";
            }
           
            SqlCommand query = connect.Query(sql);

            string portAddress = null;

            string portCode = null;

            using (SqlDataReader reader = query.ExecuteReader())
            {
                if (reader.Read())
                {
                    portAddress = reader.GetString(0).Trim() + "," + reader.GetString(1).Trim() + "," + reader.GetString(2).Trim() + "," + reader.GetString(3).Trim();

                    portCode = reader.GetString(4).Trim();
                }
            }

            if(portAddress == null)
            {
                string newSQl = $"SELECT Port_StreetAddress,Port_CityName,Port_StateCode,Port_CountryCode,Port_Code FROM [dbo].[PortAddress] Where Port_CountryCode='{address.Last().Trim()}'";

                SqlCommand newQuery = connect.Query(newSQl);

                using (SqlDataReader reader = newQuery.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        portAddress = reader.GetString(0).Trim() + "," + reader.GetString(1).Trim() + "," + reader.GetString(2).Trim() + "," + reader.GetString(3).Trim();
                        portCode = reader.GetString(4).Trim();
                    }
                }
            }

            string[] strings = { portAddress, portCode };
            return strings;
        }

        private string GetProjectAddress()
        {
            IfcSite ifcSite = Model.Instances.OfType<IfcSite>().FirstOrDefault();

            if (ifcSite == null)
            {
                return GetAddressByBuilding();

            }
            var address = ifcSite.SiteAddress;

            if (address == null)
            {
                return GetAddressByBuilding();
            }

            string projectAddress = string.Join(",", address.AddressLines.ToList());

            return projectAddress;
        }

        private string GetAddressByBuilding()
        {
            IfcBuilding ifcBuilding = Model.Instances.OfType<IfcBuilding>().FirstOrDefault();

            if (ifcBuilding == null)
            {
                return null;
            }

            var address = ifcBuilding.BuildingAddress;

            if (address == null)
            {
                return null;
            }

            string projectAddress = string.Join(",", address.AddressLines.ToList());

            return projectAddress;
        }

        #endregion

        #region GetDistanceData

        private List<(double, string, Record)> GetDistanceDataInternational(string projectAddress, string supplierAddress, string materialType)
        {
            List<(double, string, Record)> distanceData = new List<(double, string, Record)>();

            string[] port_Gate = GetPortAddress(supplierAddress); //port near supplier

            string[] port_Site = GetPortAddress(projectAddress); //port near project site

            string portAddress_Gate = port_Gate[0].Replace(",N/A", "");

            string portAddress_Site = port_Site[0].Replace(",N/A", "");

            string projectAddress_New = projectAddress.Replace(",N/A", "");

            string supplierAddress_New = supplierAddress.Replace(",N/A", "");

            if (portAddress_Gate == null || portAddress_Site == null)
            {
                MessageBox.Show("Error! No port address found");

                return null; //return null if no port address found
            }

            string[] portAddressString_Gate = portAddress_Gate.Split(',');

            string[] portAddressString_Site = portAddress_Site.Split(',');

            //double distance_G2P = 10;

            //double distance_P2S = 20;

            double distance_G2P = distanceCalculator.CalculateLandDistance(supplierAddress_New, portAddress_Gate) / 1000;

            double distance_P2S = distanceCalculator.CalculateLandDistance(projectAddress_New, portAddress_Site) / 1000;

            double distance_P2P = 2480;

            //double distance_P2P = distanceCalculator.CalculateMaritimeDistance(port_Gate[1], port_Site[1]) / 1000;

            List<string> transportModes = GetTransportModes(materialType);

            if (transportModes.Count == 0)
            {
                MessageBox.Show("Error! No avilable transport modes found");

                return null;
            }

            string[] GeoScope_G2P = { portAddressString_Gate[portAddressString_Gate.Length - 2], portAddressString_Gate.Last() };

            string[] GeoScope_P2S = { portAddressString_Site[portAddressString_Site.Length - 2], portAddressString_Site.Last()};

            Record record_G2P = GetTransportData(transportModes, GeoScope_G2P, "Domestic");

            Record record_P2S = GetTransportData(transportModes, GeoScope_P2S, "Domestic");

            string[] GeoScope_P2P = { "Global","International" };

            Record record_P2P = GetTransportData(transportModes, GeoScope_P2P, "International");

            if(record_G2P == null || record_P2S == null || record_P2P == null)
            {
                MessageBox.Show("Error! No avilable transport LCI dataset found");

                return distanceData;
            }

            distanceData.Add((distance_G2P, "km", record_G2P));

            distanceData.Add((distance_P2S, "km", record_P2S));

            distanceData.Add((distance_P2P, "km", record_P2P));

            return distanceData;
        }

        private (double, string, Record) GetDistanceDataDomestic(string projectAddress, string a, string materialType, string[] GeoScope)
        {
            //double distance = 30;
            double distance = distanceCalculator.CalculateLandDistance(projectAddress.Replace(",N/A", ""), a.Replace(",N/A", "")) / 1000;

            List<string> transportModes = GetTransportModes(materialType);

            if(transportModes.Count == 0)
            {
                MessageBox.Show("Error! No avilable transport modes found");
                return (0, "N/A", null);
            }

            Record LCARecord = GetTransportData(transportModes, GeoScope, "Domestic");

            if(LCARecord == null)
            {
                MessageBox.Show("Error! No avilable transport LCI found");
                return (0, "N/A", null);
            }

            return (distance, "km", LCARecord);
        }

        private List<string> GetTransportModes(string materialType)
        {
            string sql = $"SELECT TransportationMode FROM [dbo].[TransportLCI_Rel] WHERE MaterialType = '{materialType}';";

            SqlCommand query = connect.Query(sql);

            List<string> transportModes = new List<string>();

            using (SqlDataReader reader = query.ExecuteReader())
            {
                while (reader.Read())
                {
                    string mode = reader.GetString(0);

                    transportModes.Add(mode);
                }
            }
            return transportModes;
        }

        private Record GetTransportData(List<string> transportModes, string[] geoScope, string international)
        {
            List<Record> records = new List<Record>();

            string sql = $"SELECT  [dbo].[TransportLCI_Type].[Transport_ID],Unit,GWP,ODP,AP,EP,POCP,ADPE,ADPF,PERT,PENRT,Transport_Mode FROM [dbo].[TransportLCI_Type]" +
                $"INNER JOIN [dbo].[TransportLCI_DataSet] ON [dbo].[TransportLCI_DataSet].[Transport_ID]=[dbo].[TransportLCI_Type].[Transport_ID]" +
                $"INNER JOIN [dbo].[TransportLCI_Info] ON [dbo].[TransportLCI_Info].[Transport_ID]=[dbo].[TransportLCI_Type].[Transport_ID]" +
                $"WHERE";

            if (transportModes.Count != 0)
            {
                sql += " [dbo].[TransportLCI_Type].[Transport_Mode] IN ('" + transportModes[0] + "'";

                for (int i = 1; i < transportModes.Count; i++)
                {
                    sql += ",'" + transportModes[i] + "'";
                }

                sql += ") AND ";
            }

            sql += $"[dbo].[TransportLCI_Info].[GeoScope] In ('{geoScope[0]}','{geoScope[1]}','Global') AND [dbo].[TransportLCI_Type].[International/Domestic]='{international}';";

            SqlCommand query = connect.Query(sql);

            using (SqlDataReader reader = query.ExecuteReader())
            {
                while (reader.Read())
                {
                    Record record = new Record()
                    {
                        Uid = reader.GetGuid(0).ToString(),
                        Unit = reader.GetString(1),
                        GWP = reader.GetDouble(2),
                        ODP = reader.GetDouble(3),
                        AP = reader.GetDouble(4),
                        EP = reader.GetDouble(5),
                        POCP = reader.GetDouble(6),
                        ADPE = reader.GetDouble(7),
                        ADPF = reader.GetDouble(8),
                        PERT = reader.GetDouble(9),
                        PENRT = reader.GetDouble(10),
                        MaterialType = reader.GetString(11)
                    };
                    records.Add(record);
                }
            }

            if (records.Count == 0)
            {
                return null;
            }

            Random random = new Random();

            //get an random record

            int randomNum = random.Next(0, records.Count - 1);

            return records[randomNum];
        }

        #endregion

        #endregion

        #region A5

        private List<EleLCAResult> GetA5Results(List<List<MaterialInfo>> eleInfos)
        {

            string file_description = Model.Header.ModelViewDefinition;

            // Regular expression pattern to match the values within [] followed by ViewDefinition
            string pattern = @"ExchangeRequirement\s*\[([^[\]]*)\]";

            // Match the pattern in the line
            Match match = Regex.Match(file_description, pattern);

            string er_Name = "";

            int samplingMode = 0;

            if (byNum.IsChecked == true)
            {
                samplingMode = 1;
            }
            else if (byPct.IsChecked == true)
            {
                samplingMode = 2;
            }

            int samplingNum = int.TryParse(samplingSize.Text, out samplingNum)? samplingNum : 0;

            // Check if a match is found
            if (match.Success)
            {
                // Get the value enclosed within square brackets
                er_Name = match.Groups[1].Value;
            }

            if(er_Name == "LOD200" || er_Name == "LOD300")
            {
                //Do the calculation

                return GetA5Results_LOD200_300(samplingMode,samplingNum);

            }
            else if(er_Name == "LOD350" || er_Name == "LOD400")
            {
                return GetA5Results_LOD350_400(eleInfos);
            }
            else
            {
                MessageBox.Show("Error! No Exchange Requirement found");

                return null;
            }


        }

        private List<EleLCAResult> GetA5Results_LOD200_300(int samplingMode, int samplingNum)
        {
            var building=Model.Instances.OfType<IfcBuilding>().FirstOrDefault();

            if(building == null)
            {
                MessageBox.Show("Error! No building found");

                return null;
            }

            var buildingType_Property = building.GetPropertySingleValue("Pset_BuildingCommon", "OccupancyType");

            string buildingType = "";

            if(buildingType_Property != null)
            {
                buildingType = buildingType_Property.NominalValue.ToString();
            }

            string constructionYear = "";

            var constructionYear_Property = building.GetPropertySingleValue("Pset_BuildingCommon", "YearOfConstruction");

            if(constructionYear_Property != null)
            {
                constructionYear = constructionYear_Property.NominalValue.ToString();
            }

            int year = int.TryParse(constructionYear, out year) ? year : 0;

            string projectAddress = GetProjectAddress();

            string nation = "";

            string state = "";

            if (projectAddress!=null)
            {
                nation= projectAddress.Split(',').Last().Trim();

                state = projectAddress.Split(',')[projectAddress.Split(',').Length - 2].Trim();
            }

            string numOfStoreys = "";
            
            var numOfStoreys_Property = building.GetPropertySingleValue("Pset_BuildingCommon", "NumberOfStoreys");

            if(numOfStoreys_Property != null)
            {
                numOfStoreys = numOfStoreys_Property.NominalValue.ToString();
            }

            int storeys = int.TryParse(numOfStoreys, out storeys)? storeys : 0;

            double buildingArea = 0;

            IfcQuantityArea constructionArea = building.GetQuantity<IfcQuantityArea>("Qto_BuildingBaseQuantities", "GrossFloorArea");

            if(constructionArea == null)
            {
                constructionArea = building.GetQuantity<IfcQuantityArea>("Qto_BuildingBaseQuantities", "NetFloorArea");

                if(constructionArea == null)
                {
                    constructionArea = building.GetQuantity<IfcQuantityArea>("Qto_BuildingBaseQuantities", "FootprintArea"); 
                    
                    if(constructionArea == null)
                    {
                        string area = building.GetPropertySingleValue("Pset_BuildingCommon", "NetPlannedArea").NominalValue.ToString();

                        buildingArea = double.TryParse(area, out buildingArea) ? buildingArea : 0;
                    }
                }             
            }

            if(constructionArea != null)
            {
                buildingArea = constructionArea.AreaValue;
            }

            double buildingHeight = 0;

            IfcQuantityLength height = building.GetQuantity<IfcQuantityLength>("Qto_BuildingBaseQuantities", "Height");

            if(height!=null)
            {
                buildingHeight = height.LengthValue/1000;
            }

            string sql = $"SELECT A5_LCAGeneralInfo.[UID],Unit from A5_LCAGeneralInfo " +
                $"INNER JOIN A5_LCADataSet ON A5_LCADataSet.UID = A5_LCAGeneralInfo.UID" +
                $" Where Name !='Dummy'";

            if(buildingType != "")
            {
                sql += $" AND BuildingType IN ('{buildingType}','No Limits')";
            }

            if(year != 0)
            {
                sql += $" AND (ValidFrom <= '{year}-01-01' AND ValidUntil >= '{year}-12-31')";
            }

            if(nation != "" || state!="")
            {
                sql += $" AND ((GeoScope IN ('{nation}','No Limits') AND GeoScopeLevel='Nation') or (GeoScope IN ('{state}','No Limits') AND GeoScopeLevel='State'))";
            }

            if(storeys != 0)
            {
                sql += $" AND ([Storey Limits_LowerBound] <= '{storeys}' AND [Storey Limits_UpperBound] >= '{storeys}')";
            }

            if(buildingHeight != 0)
            {
                sql += $" AND [Height Limits] >= '{buildingHeight}'";
            }

            sql += ";";

            SqlCommand query = connect.Query(sql);

            List<(Guid,string)> UIDs = new List<(Guid,string)>();

            using (SqlDataReader reader = query.ExecuteReader())
            {
                while (reader.Read())
                {
                    UIDs.Add((reader.GetGuid(0),reader.GetString(1)));
                }
            }

            List<Guid> sampleGuids = SamplingA5Guids(UIDs, samplingNum, samplingMode);

            if(sampleGuids == null || sampleGuids.Count==0)
            {
                MessageBox.Show("Error! No sample data found");

                return null;
            }

            List<Record> records = new List<Record>();

            foreach(var guid in sampleGuids)
            {
                Record unitRecord = RetriveA5SpecificData(guid);

                Record newRecord= new Record()
                {
                    Uid = unitRecord.Uid,

                    Unit = unitRecord.Unit,

                    GWP = unitRecord.GWP * buildingArea,

                    ODP = unitRecord.ODP * buildingArea,

                    AP = unitRecord.AP * buildingArea,

                    EP = unitRecord.EP * buildingArea,

                    POCP = unitRecord.POCP * buildingArea,

                    ADPE = unitRecord.ADPE * buildingArea,

                    ADPF = unitRecord.ADPF * buildingArea,

                    PERT = unitRecord.PERT * buildingArea,

                    PENRT = unitRecord.PENRT * buildingArea
                };

                records.Add(newRecord);
            }

            List<EleLCAResult> eleResults = new List<EleLCAResult>();

            EleLCAResult eleResult = new EleLCAResult();

            eleResult.Name = building.Name.ToString();

            eleResult.IfcLabel = building.EntityLabel.ToString();

            eleResult.Stage = "A5";

            eleResult.Description = $"The calculation is for {eleResult.Name} in total, {records.Count} number of records are sampled";

            eleResult.materialCount = 0;

            eleResult.MaterialResults = new Dictionary<string, List<Record>>();

            eleResult.MaterialResults.Add(eleResult.Name, records);

            CalculateStaticticResults(eleResult);

            eleResults.Add(eleResult);

            return eleResults;
        }

        private Record RetriveA5SpecificData(Guid guid)
        {
            string sql=$"SELECT * FROM A5_LCADataSet WHERE UID='{guid}';";

            SqlCommand query = connect.Query(sql);

            Record record = new Record();

            using (SqlDataReader reader = query.ExecuteReader())
            {
                if (reader.Read())
                {
                    record.Uid = reader.GetGuid(0).ToString();

                    record.Unit = reader.GetString(1);

                    record.GWP = reader.GetDouble(2);

                    record.ODP = reader.GetDouble(3);

                    record.AP = reader.GetDouble(4);

                    record.EP = reader.GetDouble(5);

                    record.POCP = reader.GetDouble(6);

                    record.ADPE = reader.GetDouble(7);

                    record.ADPF = reader.GetDouble(8);

                    record.PERT = reader.GetDouble(9);

                    record.PENRT = reader.GetDouble(10);

                }
            }

            return record;
        }

        private List<Guid> SamplingA5Guids (List<(Guid, string)> data, int samplingNum, int samplingMode)
        {
            List<Guid> guids = new List<Guid>();

            foreach (var pair in data)
            {
                if(DetermineUnitCategory(pair.Item2) == "Area")
                {
                    guids.Add(pair.Item1);
                }
            }

            int count = guids.Count;

            if (samplingMode == 1)
            {
                if (count <= samplingNum)
                {
                    return guids;
                }

                Random random = new Random();

                List<Guid> sampleGuids = new List<Guid>();

                while (sampleGuids.Count < samplingNum)
                {
                    int index = random.Next(0, count);

                    if (!sampleGuids.Contains(guids[index]))
                    {
                        sampleGuids.Add(guids[index]);
                    }
                }
                return sampleGuids;
            }
            else if (samplingMode == 2)
            {
                int sampleSize = (int)Math.Ceiling(count * samplingNum / 100.0);

                if (count <= sampleSize)
                {
                    return guids;
                }

                Random random = new Random();

                List<Guid> sampleGuids = new List<Guid>();

                while (sampleGuids.Count < sampleSize)
                {
                    int index = random.Next(0, count);

                    if (!sampleGuids.Contains(guids[index]))
                    {
                        sampleGuids.Add(guids[index]);
                    }
                }
                return sampleGuids;
            }

            return null;
        }

        private List<EleLCAResult> GetA5Results_LOD350_400(List<List<MaterialInfo>> eleInfos)
        {
            EleLCAResult eleResult = GetA5Results_LOD200_300(1, 1)[0];

            Record averageRecord = eleResult.TotalResult_Average;

            List<EleLCAResult> results = new List<EleLCAResult>();

            foreach (var ele in SelectedElements)
            {
                if(ele is IfcCovering || ele is IfcRoof || ele is IfcColumn || ele is IfcSlab || ele is IfcWall || ele is IfcDoor ||ele is IfcWindow || ele is IfcChimney)
                {
                    int i = ele.EntityLabel;
                    Record record = new Record()
                    {
                        GWP = averageRecord.GWP * aT(i), 
                        ODP = averageRecord.ODP * aT(i),
                        AP = averageRecord.AP * aT(i),
                        EP = averageRecord.EP * aT(i),
                        POCP = averageRecord.POCP * aT(i),
                        ADPE = averageRecord.ADPE * aT(i),
                        ADPF = averageRecord.ADPF * aT(i),
                        PENRT = averageRecord.PENRT * aT(i),
                        PERT = averageRecord.PERT * aT(i)
                    };

                    EleLCAResult eleLCAResult = new EleLCAResult();

                    eleLCAResult.Name = ele.Name.ToString();

                    eleLCAResult.Description = $"A5 Results for {eleLCAResult.Name}";

                    eleLCAResult.IfcLabel = ele.EntityLabel.ToString();

                    eleLCAResult.Stage = "A5";

                    eleLCAResult.materialCount = 0;

                    eleLCAResult.MaterialResults = new Dictionary<string, List<Record>>();

                    List<Record> records = new List<Record>();

                    records.Add(record);

                    eleLCAResult.MaterialResults.Add(eleLCAResult.Name, records);

                    CalculateStaticticResults(eleLCAResult);

                    results.Add(eleLCAResult);

                }
                else
                {
                    EleLCAResult eleLCAResult = new EleLCAResult();

                    eleLCAResult.Name = ele.Name.ToString();

                    eleLCAResult.Description = $"A5 stage not applicable";

                    eleLCAResult.IfcLabel = ele.EntityLabel.ToString();

                    eleLCAResult.Stage = "A5";

                    eleLCAResult.materialCount = 0;

                    Record record = new Record()
                    {
                        GWP = 0,
                        ODP = 0,
                        AP = 0,
                        EP = 0,
                        POCP = 0,
                        ADPE = 0,
                        ADPF = 0,
                        PENRT = 0,
                        PERT = 0
                    };

                    eleLCAResult.MaterialResults = new Dictionary<string, List<Record>>();

                    List<Record> records = new List<Record>();

                    records.Add(record);

                    eleLCAResult.MaterialResults.Add(eleLCAResult.Name, records);

                    CalculateStaticticResults(eleLCAResult);

                    results.Add(eleLCAResult);
                }
            }

            return results;
        }

        #endregion

        #region GetQuantities

        private List<List<MaterialInfo>> GetQuantities(List<List<MaterialInfo>> eleInfos)
        {
            foreach (var ele in eleInfos)
            {
                foreach (var material in ele)
                {
                    (double, string) quantity = (0, "N/A");

                    if (material.MaterialRecords.Count == 0)
                    {
                        material.quantity = 0;

                        material.unit = "N/A";

                        continue;
                    }

                    string UnitCatrgory = DetermineUnitCategory(material.MaterialRecords[0].Unit);

                    string[] stringArray = material.ID.Split('-');

                    string densityString = stringArray.Last();

                    string thicknessString = stringArray[stringArray.Length - 2];

                    double density = 0;

                    double thickness = 0;

                    bool t_S = double.TryParse(thicknessString, out thickness);

                    bool d_S = double.TryParse(densityString, out density);

                    quantity = GetQuantityByUnit(UnitCatrgory, material.IfcLabel, material.IfcType, thickness, density);
                    
                    material.quantity = quantity.Item1;

                    material.unit = quantity.Item2;
                }
            }

            return eleInfos;
        }

        private (double, string) GetQuantityByUnit(string unitCategory, string ifcLabel, string ifcType, double thickness, double density)
        {
            IfcElement ifcElement = Model.Instances.OfType<IfcElement>().FirstOrDefault(x => x.EntityLabel.ToString() == ifcLabel);

            switch (ifcType)
            {
                case "IfcDoor":
                    IfcDoor door = ifcElement as IfcDoor;
                    return GetDoorQuantity(door, unitCategory);
                case "IfcWindow":
                    IfcWindow window = ifcElement as IfcWindow;
                    return GetWindowQuantity(window, unitCategory);
                case "IfcWall":
                    IfcWall wall = ifcElement as IfcWall;
                    return GetWallQuantity(wall, unitCategory, thickness, density);
                case "IfcCurtainWall":
                    IfcCurtainWall curtainWall = ifcElement as IfcCurtainWall;
                    return GetCurtainWallQuantity(curtainWall, unitCategory, thickness, density);
                case "IfcSlab":
                    IfcSlab slab = ifcElement as IfcSlab;
                    return GetSlabQuantity(slab, unitCategory, thickness, density);
                case "IfcRoof":
                    IfcRoof roof = ifcElement as IfcRoof;
                    return GetRoofQuantity(roof, unitCategory, thickness, density);
                case "IfcCovering":
                    IfcCovering covering = ifcElement as IfcCovering;
                    return GetCoveringQuantity(covering, unitCategory, thickness, density);
                case "IfcStairFlight":
                    IfcStairFlight stairFlight = ifcElement as IfcStairFlight;
                    return GetStairFlightQuantity(stairFlight, unitCategory, density);
                case "IfcRampFlight":
                    IfcRampFlight rampFlight = ifcElement as IfcRampFlight;
                    return GetRampFlightQuantity(rampFlight, unitCategory, density);
                case "IfcRailing":
                    IfcRailing railing = ifcElement as IfcRailing;
                    return GetRailingQuantity(railing, unitCategory, density);
                case "IfcBeam":
                    IfcBeam beam = ifcElement as IfcBeam;
                    return GetBeamQuantity(beam, unitCategory, density);
                case "IfcColumn":
                    IfcColumn column = ifcElement as IfcColumn;
                    return GetColumnQuantity(column, unitCategory, density);
                case "IfcFooting":
                    IfcFooting footing = ifcElement as IfcFooting;
                    return GetFootingQuantity(footing, unitCategory, density);
                case "IfcPile":
                    IfcPile pile = ifcElement as IfcPile;
                    return GetPileQuantity(pile, unitCategory, density);
                case "IfcChimney":
                    IfcChimney chimney = ifcElement as IfcChimney;
                    return GetChimneyQuantity(chimney, unitCategory, density);
                case "IfcBuildingElementProxy":
                    IfcBuildingElementProxy proxy = ifcElement as IfcBuildingElementProxy;
                    return GetBuildingElementProxyQuantity(proxy, unitCategory, density);
                case "IfcMember":
                    IfcMember member = ifcElement as IfcMember;
                    return GetMemberQuantity(member, unitCategory, density);
                case "IfcPlate":
                    IfcPlate plate = ifcElement as IfcPlate;
                    return GetPlateQuantity(plate, unitCategory, thickness, density);
                case "IfcFurniture":
                    IfcFurniture furniture = ifcElement as IfcFurniture;
                    return GetFurnitureQuantity(furniture, unitCategory);
                case "IfcReinforcingBar":
                        IfcReinforcingBar reinforcingBar = ifcElement as IfcReinforcingBar;
                    return GetReinforcingBarQuantity(reinforcingBar, unitCategory, density);
                default:
                    return (0, "N/A");
            }
        }

        private (double, string) GetReinforcingBarQuantity(IfcReinforcingBar reinforcingBar, string unitCategory, double density)
        {
            if(unitCategory=="Length")
            {
                var lengthQuantity = reinforcingBar.PropertySets.OfType<IfcPropertySet>().FirstOrDefault(x => x.Name == "Pset_ReinforcingBarCommon")?.HasProperties.OfType<IfcPropertySingleValue>().FirstOrDefault(x => x.Name == "BarLength");

                double length = (double)lengthQuantity.NominalValue.Value;

                return ((double)length/1000, "m");
            }
            else if(unitCategory=="Volume")
            {
                var lengthQuantity=GetReinforcingBarQuantity(reinforcingBar, "Length", density);

                if(lengthQuantity.Item2=="N/A")
                {
                    return (0, "N/A");
                }

                var reference= reinforcingBar.PropertySets.OfType<IfcPropertySet>().FirstOrDefault(x => x.Name == "Pset_ReinforcingBarCommon")?.HasProperties.OfType<IfcPropertySingleValue>().FirstOrDefault(x => x.Name == "Reference");

                var reference_Text = (string)reference.NominalValue.Value;

                double diameter = 0;

                if(reference_Text!=null)
                {
                    string[] text= reference_Text.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                    diameter= double.Parse(text[0])/1000;
                }

                double volume = lengthQuantity.Item1 * Math.PI * Math.Pow(diameter / 2, 2);

                return (volume, "m3");
            }
            else if(unitCategory=="Weight")
            {
                var weight =GetReinforcingBarQuantity(reinforcingBar, "Volume", density);

                if(weight.Item2=="N/A")
                {
                    return (0, "N/A");
                }

                return (weight.Item1 * density/1000, "ton");
            }
            else if(unitCategory==".ea")
            {
                return (1, ".ea");
            }

            return (0, "N/A");
        }

        private (double, string) GetFurnitureQuantity(IfcFurniture furniture, string unitCategory)
        {
            if (unitCategory == ".ea")
            {
                return (1, ".ea");
            }

            return (0, "N/A");
        }//L

        private (double, string) GetRailingQuantity(IfcRailing railing, string unitCategory, double density)
        {
            if (unitCategory == "Length")
            {
                var lengthQuantity = railing.GetQuantity<IfcQuantityLength>("Qto_RailingBaseQuantities", "Length");

                if (lengthQuantity == null)
                {
                    failedMaterials.Add((railing.EntityLabel.ToString(), "Qto_RailingBaseQuantities-Length"));

                    return (0, "N/A");
                }

                string unit = "mm";

                if (lengthQuantity.Unit != null)
                {
                    unit = lengthQuantity.Unit.ToString();
                }

                return (lengthQuantity.LengthValue, unit);
            }
            else if (unitCategory == "Volume")
            {
                var volumeQuantity = railing.GetQuantity<IfcQuantityVolume>("Cpset_RailingAddQuantities", "NetVolume");

                if (volumeQuantity == null)
                {
                    failedMaterials.Add((railing.EntityLabel.ToString(), "Cpset_RailingAddQuantities-NetVolume"));

                    volumeQuantity = railing.GetQuantity<IfcQuantityVolume>("Cpset_RailingAddQuantities", "GrossVolume");

                    if (volumeQuantity == null)
                    {
                        failedMaterials.Add((railing.EntityLabel.ToString(), "Cpset_RailingAddQuantities-GrossVolume"));

                        return (0, "N/A");
                    }
                }

                double volume = volumeQuantity.VolumeValue;

                string unit = "m3";

                if (volumeQuantity.Unit != null)
                {
                    unit = volumeQuantity.Unit.ToString();
                }

                return (volume, unit);
            }
            else if (unitCategory == "Weight")
            {
                double weight = 0;

                var weightQuantity = railing.GetQuantity<IfcQuantityWeight>("Cpset_RailingAddQuantities", "NetWeight");

                if (weightQuantity == null)
                {
                    failedMaterials.Add((railing.EntityLabel.ToString(), "Cpset_RailingAddQuantities-NetWeight"));

                    weightQuantity = railing.GetQuantity<IfcQuantityWeight>("Cpset_RailingAddQuantities", "GrossWeight");

                    if (weightQuantity == null)
                    {
                        failedMaterials.Add((railing.EntityLabel.ToString(), "Cpset_RailingAddQuantities-GrossWeight"));

                        if (density != 0)
                        {
                            (double, string) volumeQuantity = GetRailingQuantity(railing, "Volume", density);

                            if (volumeQuantity.Item2 == "N/A")
                            {
                                return (0, "N/A");
                            }

                            UnitConverter unitConverter = new UnitConverter(volumeQuantity.Item2, "m3");

                            double volume = unitConverter.Convert(volumeQuantity.Item1);

                            weight = volume * density;

                            return (weight, "kg");
                        }

                        return (0, "N/A");
                    }
                }

                weight = weightQuantity.WeightValue;

                string unit = "kg";

                if (weightQuantity.Unit != null)
                {
                    unit = weightQuantity.Unit.ToString();
                }

                return (weight, unit);
            }
            else if (unitCategory == ".ea")
            {
                return (1, ".ea");
            }

            return (0, "N/A");
        }//L

        private (double, string) GetPlateQuantity(IfcPlate plate, string unitCategory, double thickness, double density)
        {
            if (unitCategory == "Area")
            {
                var areaQuantity = plate.GetQuantity<IfcQuantityArea>("Qto_PlateBaseQuantities", "NetArea");

                if (areaQuantity == null)
                {
                    failedMaterials.Add((plate.EntityLabel.ToString(), "Qto_PlateBaseQuantities-NetArea"));

                    areaQuantity = plate.GetQuantity<IfcQuantityArea>("Qto_PlateBaseQuantities", "GrossArea");

                    if (areaQuantity == null)
                    {
                        failedMaterials.Add((plate.EntityLabel.ToString(), "Qto_PlateBaseQuantities-GrossArea"));

                        return (0, "N/A");
                    }
                }

                double area = areaQuantity.AreaValue;

                string unit = "m2";

                if (areaQuantity.Unit != null)
                {
                    unit = areaQuantity.Unit.ToString();
                }

                return (area, unit);
            }
            else if (unitCategory == "Volume")
            {
                var volumeQuantity = plate.GetQuantity<IfcQuantityVolume>("Qto_PlateBaseQuantities", "NetVolume");

                if (volumeQuantity == null)
                {
                    failedMaterials.Add((plate.EntityLabel.ToString(), "Qto_PlateBaseQuantities-NetVolume"));

                    volumeQuantity = plate.GetQuantity<IfcQuantityVolume>("Qto_PlateBaseQuantities", "GrossVolume");

                    if (volumeQuantity == null)
                    {
                        failedMaterials.Add((plate.EntityLabel.ToString(), "Qto_PlateBaseQuantities-GrossVolume"));

                        var areaQuanity = GetPlateQuantity(plate, "Area", thickness, density);

                        if (areaQuanity.Item2 == "N/A")
                        {
                            return (0, "N/A");
                        }

                        var depth = plate.GetQuantity<IfcQuantityLength>("Qto_PlateBaseQuantities", "Width");

                        UnitConverter unitConverter = new UnitConverter(areaQuanity.Item2, "m2");

                        double area = unitConverter.Convert(areaQuanity.Item1);

                        if (depth == null)
                        {
                            failedMaterials.Add((plate.EntityLabel.ToString(), "Qto_PlateBaseQuantities-Width"));

                            return (thickness * areaQuanity.Item1 / 1000, "m3");
                        }
                        else
                        {
                            unitConverter = new UnitConverter(depth.Unit.Name(), "m");

                            double depthValue = unitConverter.Convert(depth.LengthValue);

                            return (depthValue * area, "m3");
                        }
                    }
                }

                double volume = volumeQuantity.VolumeValue;

                string unit = "m3";

                if (volumeQuantity.Unit != null)
                {
                    unit = volumeQuantity.Unit.ToString();
                }

                return (volume, unit);
            }
            else if (unitCategory == "Weight")
            {
                double weight = 0;

                var weightQuantity = plate.GetQuantity<IfcQuantityWeight>("Qto_PlateBaseQuantities", "NetWeight");

                if (weightQuantity == null)
                {
                    failedMaterials.Add((plate.EntityLabel.ToString(), "Qto_PlateBaseQuantities-NetWeight"));

                    weightQuantity = plate.GetQuantity<IfcQuantityWeight>("Qto_PlateBaseQuantities", "GrossWeight");

                    if (weightQuantity == null)
                    {
                        failedMaterials.Add((plate.EntityLabel.ToString(), "Qto_PlateBaseQuantities-GrossWeight"));

                        if (density != 0)
                        {
                            (double, string) volumeQuantity = GetPlateQuantity(plate, "Volume", thickness, density);

                            if (volumeQuantity.Item2 == "N/A")
                            {
                                return (0, "N/A");
                            }

                            UnitConverter unitConverter = new UnitConverter(volumeQuantity.Item2, "m3");

                            double volume = unitConverter.Convert(volumeQuantity.Item1);

                            weight = volume * density;

                            return (weight, "kg");
                        }

                        return (0, "N/A");
                    }
                }

                weight = weightQuantity.WeightValue;

                string unit = "kg";

                if (weightQuantity.Unit != null)
                {
                    unit = weightQuantity.Unit.ToString();
                }

                return (weight, unit);
            }
            else if (unitCategory == ".ea")
            {
                return (1, ".ea");
            }

            return (0, "N/A");
        }

        private (double, string) GetMemberQuantity(IfcMember member, string unitCategory, double density)
        {
            if (unitCategory == "Length")
            {
                var lengthQuantity = member.GetQuantity<IfcQuantityLength>("Qto_MemberBaseQuantities", "Length");

                if (lengthQuantity == null)
                {
                    failedMaterials.Add((member.EntityLabel.ToString(), "Qto_MemberBaseQuantities-Length"));

                    return (0, "N/A");
                }

                string unit = "mm";

                if (lengthQuantity.Unit != null)
                {
                    unit = lengthQuantity.Unit.ToString();
                }

                return (lengthQuantity.LengthValue, unit);
            }
            else if (unitCategory == "Volume")
            {
                var volumeQuantity = member.GetQuantity<IfcQuantityVolume>("Qto_MemberBaseQuantities", "NetVolume");

                if (volumeQuantity == null)
                {
                    failedMaterials.Add((member.EntityLabel.ToString(), "Qto_MemberBaseQuantities-NetVolume"));

                    volumeQuantity = member.GetQuantity<IfcQuantityVolume>("Qto_MemberBaseQuantities", "GrossVolume");

                    if (volumeQuantity == null)
                    {
                        failedMaterials.Add((member.EntityLabel.ToString(), "Qto_MemberBaseQuantities-GrossVolume"));

                        var lengthQuantity = member.GetQuantity<IfcQuantityLength>("Qto_MemberBaseQuantities", "Length");

                        var areaQuantity = member.GetQuantity<IfcQuantityArea>("Qto_MemberBaseQuantities", "CrossSectionArea");

                        if (lengthQuantity != null && areaQuantity != null)
                        {
                            UnitConverter unitConverter = new UnitConverter(lengthQuantity.Unit.Name(), "m");

                            double length = unitConverter.Convert(lengthQuantity.LengthValue);

                            unitConverter = new UnitConverter(areaQuantity.Unit.Name(), "m2");

                            double area = unitConverter.Convert(areaQuantity.AreaValue);

                            return (length * area, "m3");
                        }
                        else
                        {
                            failedMaterials.Add((member.EntityLabel.ToString(), "Qto_MemberBaseQuantities-Length"));

                            failedMaterials.Add((member.EntityLabel.ToString(), "Qto_MemberBaseQuantities-CrossSectionArea"));

                            return (0, "N/A");
                        }
                    }
                }

                double volume = volumeQuantity.VolumeValue;

                string unit = "m3";

                if (volumeQuantity.Unit != null)
                {
                    unit = volumeQuantity.Unit.ToString();
                }

                return (volume, unit);
            }
            else if (unitCategory == "Weight")
            {
                double weight = 0;

                var weightQuantity = member.GetQuantity<IfcQuantityWeight>("Qto_MemberBaseQuantities", "NetWeight");

                if (weightQuantity == null)
                {
                    failedMaterials.Add((member.EntityLabel.ToString(), "Qto_MemberBaseQuantities-NetWeight"));

                    weightQuantity = member.GetQuantity<IfcQuantityWeight>("Qto_MemberBaseQuantities", "GrossWeight");

                    if (weightQuantity == null)
                    {
                        failedMaterials.Add((member.EntityLabel.ToString(), "Qto_MemberBaseQuantities-GrossWeight"));

                        if (density != 0)
                        {
                            (double, string) volumeQuantity = GetMemberQuantity(member, "Volume", density);

                            if (volumeQuantity.Item2 == "N/A")
                            {
                                return (0, "N/A");
                            }

                            UnitConverter unitConverter = new UnitConverter(volumeQuantity.Item2, "m3");

                            double volume = unitConverter.Convert(volumeQuantity.Item1);

                            weight = volume * density;

                            return (weight, "kg");
                        }

                        return (0, "N/A");
                    }
                }

                weight = weightQuantity.WeightValue;

                string unit = "kg";

                if (weightQuantity.Unit != null)
                {
                    unit = weightQuantity.Unit.ToString();
                }

                return (weight, unit);
            }
            else if (unitCategory == ".ea")
            {
                return (1, ".ea");
            }

            return (0, "N/A");
        }//L

        private (double, string) GetBuildingElementProxyQuantity(IfcBuildingElementProxy proxy, string unitCategory, double density)
        {
            if (unitCategory == "Volume")
            {
                var volumeQuantity = proxy.GetQuantity<IfcQuantityVolume>("Qto_BuildingElementProxyQuantities", "NetVolume");

                if (volumeQuantity == null)
                {
                    failedMaterials.Add((proxy.EntityLabel.ToString(), "Qto_BuildingElementProxyQuantities-NetVolume"));

                    return (0, "N/A");
                }

                double volume = volumeQuantity.VolumeValue;

                string unit = "m3";

                if (volumeQuantity.Unit != null)
                {
                    volumeQuantity.Unit.ToString();
                }

                return (volume, unit);
            }
            else if (unitCategory == "Weight")
            {
                if (density != 0)
                {
                    (double, string) volumeQuantity = GetBuildingElementProxyQuantity(proxy, "Volume", density);

                    if (volumeQuantity.Item2 == "N/A")
                    {
                        return (0, "N/A");
                    }

                    UnitConverter unitConverter = new UnitConverter(volumeQuantity.Item2, "m3");

                    double volume = unitConverter.Convert(volumeQuantity.Item1);

                    double weight = volume * density;

                    return (weight, "kg");
                }

                return (0, "N/A");

            }
            else if (unitCategory == ".ea")
            {
                return (1, ".ea");
            }

            return (0, "N/A");
        }//L

        private (double, string) GetChimneyQuantity(IfcChimney chimney, string unitCategory, double density)
        {
            if (unitCategory == "Length")
            {
                var lengthQuantity = chimney.GetQuantity<IfcQuantityLength>("Cpset_ChimneyAddQuantities", "Length");

                if (lengthQuantity == null)
                {
                    failedMaterials.Add((chimney.EntityLabel.ToString(), "Cpset_ChimneyAddQuantities-Length"));

                    return (0, "N/A");
                }

                string unit = "mm";

                if (lengthQuantity.Unit != null)
                {
                    unit = lengthQuantity.Unit.ToString();
                }

                return (lengthQuantity.LengthValue, unit);
            }
            else if (unitCategory == "Volume")
            {
                //var volumeQuantity = chimney.GetQuantity<IfcQuantityVolume>("Cpset_ChimneyAddQuantities", "NetVolume");

                //var volume=chimney.GetPropertySingleValue("Pset_ChimneyAddQuantities", "NetVolume");

                var volume = chimney.IsDefinedBy.OfType<IfcRelDefinesByProperties>().SelectMany(r => r.RelatingPropertyDefinition.PropertySetDefinitions).OfType<IfcPropertySet>().SelectMany(p => p.HasProperties).OfType<IfcPropertySingleValue>().FirstOrDefault(p => p.Name == "NetVolume");

                if (volume == null)
                {
                    failedMaterials.Add((chimney.EntityLabel.ToString(), "Cpset_ChimneyAddQuantities-NetVolume"));

                    //volumeQuantity = chimney.GetQuantity<IfcQuantityVolume>("Cpset_ChimneyAddQuantities", "GrossVolume");

                    volume = chimney.IsDefinedBy.OfType<IfcRelDefinesByProperties>().SelectMany(r => r.RelatingPropertyDefinition.PropertySetDefinitions).OfType<IfcPropertySet>().SelectMany(p => p.HasProperties).OfType<IfcPropertySingleValue>().FirstOrDefault(p => p.Name == "GrossVolume");

                    if (volume == null)
                    {
                        failedMaterials.Add((chimney.EntityLabel.ToString(), "Cpset_ChimneyAddQuantities-GrossVolume"));

                        return (0, "N/A");
                    }
                }

                double volumeQuantity = 0;
                    
                double.TryParse(volume.NominalValue.ToString(), out volumeQuantity);

                return (volumeQuantity, "m3");
            }
            else if (unitCategory == "Weight")
            {
                double weight = 0;

                var weightQuantity = chimney.GetQuantity<IfcQuantityWeight>("Cpset_ChimneyAddQuantities", "NetWeight");

                if (weightQuantity == null)
                {
                    failedMaterials.Add((chimney.EntityLabel.ToString(), "Cpset_ChimneyAddQuantities-NetWeight"));

                    weightQuantity = chimney.GetQuantity<IfcQuantityWeight>("Cpset_ChimneyAddQuantities", "GrossWeight");

                    if (weightQuantity == null)
                    {
                        failedMaterials.Add((chimney.EntityLabel.ToString(), "Cpset_ChimneyAddQuantities-GrossWeight"));

                        if (density != 0)
                        {
                            (double, string) volumeQuantity = GetChimneyQuantity(chimney, "Volume", density);

                            if (volumeQuantity.Item2 == "N/A")
                            {
                                return (0, "N/A");
                            }

                            UnitConverter unitConverter = new UnitConverter(volumeQuantity.Item2, "m3");

                            double volume = unitConverter.Convert(volumeQuantity.Item1);

                            weight = volume * density;

                            return (weight, "kg");
                        }

                        return (0, "N/A");
                    }
                }

                weight = weightQuantity.WeightValue;

                string unit = "kg";

                if (weightQuantity.Unit != null)
                {
                    unit = weightQuantity.Unit.ToString();
                }

                return (weight, unit);
            }
            else if (unitCategory == ".ea")
            {
                return (1, ".ea");
            }

            return (0, "N/A");
        }//L

        private (double, string) GetPileQuantity(IfcPile pile, string unitCategory, double density)
        {
            if (unitCategory == "Length")
            {
                var lengthQuantity = pile.GetQuantity<IfcQuantityLength>("Qto_PileBaseQuantities", "Length");

                if (lengthQuantity == null)
                {
                    failedMaterials.Add((pile.EntityLabel.ToString(), "Qto_PileBaseQuantities-Length"));

                    return (0, "N/A");
                }

                string unit = "mm";

                if (lengthQuantity.Unit != null)
                {
                    unit = lengthQuantity.Unit.ToString();
                }   

                return (lengthQuantity.LengthValue, unit);
            }
            else if (unitCategory == "Volume")
            {
                var volumeQuantity = pile.GetQuantity<IfcQuantityVolume>("Qto_PileBaseQuantities", "NetVolume");

                if (volumeQuantity == null)
                {
                    failedMaterials.Add((pile.EntityLabel.ToString(), "Qto_PileBaseQuantities-NetVolume"));

                    volumeQuantity = pile.GetQuantity<IfcQuantityVolume>("Qto_PileBaseQuantities", "GrossVolume");

                    if (volumeQuantity == null)
                    {
                        failedMaterials.Add((pile.EntityLabel.ToString(), "Qto_PileBaseQuantities-GrossVolume"));

                        var lengthQuantity = pile.GetQuantity<IfcQuantityLength>("Qto_PileBaseQuantities", "Length");

                        var areaQuantity = pile.GetQuantity<IfcQuantityArea>("Qto_PileBaseQuantities", "CrossSectionArea");

                        if (lengthQuantity != null && areaQuantity != null)
                        {
                            UnitConverter unitConverter = new UnitConverter(lengthQuantity.Unit.Name(), "m");

                            double length = unitConverter.Convert(lengthQuantity.LengthValue);

                            unitConverter = new UnitConverter(areaQuantity.Unit.Name(), "m2");

                            double area = unitConverter.Convert(areaQuantity.AreaValue);

                            return (length * area, "m3");
                        }
                        else
                        {
                            failedMaterials.Add((pile.EntityLabel.ToString(), "Qto_PileBaseQuantities-Length"));

                            failedMaterials.Add((pile.EntityLabel.ToString(), "Qto_PileBaseQuantities-CrossSectionArea"));

                            return (0, "N/A");
                        }
                    }
                }

                double volume = volumeQuantity.VolumeValue;

                string unit = "m3";

                if (volumeQuantity.Unit != null)
                {
                    unit = volumeQuantity.Unit.ToString();
                }

                return (volume, unit);
            }
            else if (unitCategory == "Weight")
            {
                double weight = 0;

                var weightQuantity = pile.GetQuantity<IfcQuantityWeight>("Qto_PileBaseQuantities", "NetWeight");

                if (weightQuantity == null)
                {
                    failedMaterials.Add((pile.EntityLabel.ToString(), "Qto_PileBaseQuantities-NetWeight"));

                    weightQuantity = pile.GetQuantity<IfcQuantityWeight>("Qto_PileBaseQuantities", "GrossWeight");

                    if (weightQuantity == null)
                    {
                        failedMaterials.Add((pile.EntityLabel.ToString(), "Qto_PileBaseQuantities-GrossWeight"));

                        if (density != 0)
                        {
                            (double, string) volumeQuantity = GetPileQuantity(pile, "Volume", density);

                            if (volumeQuantity.Item2 == "N/A")
                            {
                                return (0, "N/A");
                            }

                            UnitConverter unitConverter = new UnitConverter(volumeQuantity.Item2, "m3");

                            double volume = unitConverter.Convert(volumeQuantity.Item1);

                            weight = volume * density;

                            return (weight, "kg");
                        }

                        return (0, "N/A");
                    }
                }

                weight = weightQuantity.WeightValue;

                string unit = "kg";

                if (weightQuantity.Unit != null)
                {
                    unit = weightQuantity.Unit.ToString();
                }

                return (weight, unit);
            }
            else if (unitCategory == ".ea")
            {
                return (1, ".ea");
            }

            return (0, "N/A");
        }//L

        private (double, string) GetFootingQuantity(IfcFooting footing, string unitCategory, double density)
        {
            if (unitCategory == "Length")
            {
                var lengthQuantity = footing.GetQuantity<IfcQuantityLength>("Qto_FootingBaseQuantities", "Length");

                if (lengthQuantity == null)
                {
                    failedMaterials.Add((footing.EntityLabel.ToString(), "Qto_FootingBaseQuantities-Length"));

                    return (0, "N/A");
                }

                string unit = "mm";

                if (lengthQuantity.Unit != null)
                {
                    unit = lengthQuantity.Unit.ToString();
                }

                return (lengthQuantity.LengthValue, unit);
            }
            else if (unitCategory == "Area")
            {
                var lengthQuantity = footing.GetQuantity<IfcQuantityLength>("Qto_FootingBaseQuantities", "Length");

                var widthQuantity = footing.GetQuantity<IfcQuantityLength>("Qto_FootingBaseQuantities", "Width");

                if (lengthQuantity != null && widthQuantity != null)
                {
                    UnitConverter unitConverter = new UnitConverter(lengthQuantity.Unit.Name(), "m");

                    double length = unitConverter.Convert(lengthQuantity.LengthValue);

                    unitConverter = new UnitConverter(widthQuantity.Unit.Name(), "m");

                    double width = unitConverter.Convert(widthQuantity.LengthValue);

                    return (length * width, "m2");
                }
                else
                {
                    failedMaterials.Add((footing.EntityLabel.ToString(), "Qto_FootingBaseQuantities-Length"));

                    failedMaterials.Add((footing.EntityLabel.ToString(), "Qto_FootingBaseQuantities-Width"));

                    return (0, "N/A");
                }
            }
            else if (unitCategory == "Volume")
            {
                var volumeQuantity = footing.GetQuantity<IfcQuantityVolume>("Qto_FootingBaseQuantities", "NetVolume");

                if (volumeQuantity == null)
                {
                    failedMaterials.Add((footing.EntityLabel.ToString(), "Qto_FootingBaseQuantities-NetVolume"));

                    volumeQuantity = footing.GetQuantity<IfcQuantityVolume>("Qto_FootingBaseQuantities", "GrossVolume");

                    if (volumeQuantity == null)
                    {
                        failedMaterials.Add((footing.EntityLabel.ToString(), "Qto_FootingBaseQuantities-GrossVolume"));

                        var lengthQuantity = footing.GetQuantity<IfcQuantityLength>("Qto_FootingBaseQuantities", "Length");

                        var areaQuantity = footing.GetQuantity<IfcQuantityArea>("Qto_FootingBaseQuantities", "CrossSectionArea");

                        if (lengthQuantity != null)
                        {
                            if (areaQuantity != null)
                            {
                                UnitConverter unitConverter = new UnitConverter(lengthQuantity.Unit.Name(), "m");

                                double length = unitConverter.Convert(lengthQuantity.LengthValue);

                                unitConverter = new UnitConverter(areaQuantity.Unit.Name(), "m2");

                                double area = unitConverter.Convert(areaQuantity.AreaValue);

                                return (length * area, "m3");
                            }
                            else
                            {
                                failedMaterials.Add((footing.EntityLabel.ToString(), "Qto_FootingBaseQuantities-CrossSectionArea"));

                                var widthQuantity = footing.GetQuantity<IfcQuantityLength>("Qto_FootingBaseQuantities", "Width");

                                var heightQuantity = footing.GetQuantity<IfcQuantityLength>("Qto_FootingBaseQuantities", "Height");

                                if (widthQuantity != null && heightQuantity != null)
                                {
                                    UnitConverter unitConverter = new UnitConverter(lengthQuantity.Unit.Name(), "m");

                                    double length = unitConverter.Convert(lengthQuantity.LengthValue);

                                    unitConverter = new UnitConverter(widthQuantity.Unit.Name(), "m");

                                    double width = unitConverter.Convert(widthQuantity.LengthValue);

                                    unitConverter = new UnitConverter(heightQuantity.Unit.Name(), "m");

                                    double height = unitConverter.Convert(heightQuantity.LengthValue);

                                    return (length * width * height, "m3");
                                }
                                else
                                {
                                    failedMaterials.Add((footing.EntityLabel.ToString(), "Qto_FootingBaseQuantities-Width"));

                                    failedMaterials.Add((footing.EntityLabel.ToString(), "Qto_FootingBaseQuantities-Height"));

                                    return (0, "N/A");
                                }
                            }
                        }
                        else
                        {
                            failedMaterials.Add((footing.EntityLabel.ToString(), "Qto_FootingBaseQuantities-Length"));

                            return (0, "N/A");
                        }
                    }
                }

                double volume = volumeQuantity.VolumeValue;

                string unit = "m3";

                if (volumeQuantity.Unit != null)
                {
                    unit = volumeQuantity.Unit.ToString();
                }

                return (volume, unit);
            }
            else if (unitCategory == "Weight")
            {
                double weight = 0;

                var weightQuantity = footing.GetQuantity<IfcQuantityWeight>("Qto_FootingBaseQuantities", "NetWeight");

                if (weightQuantity == null)
                {
                    failedMaterials.Add((footing.EntityLabel.ToString(), "Qto_FootingBaseQuantities-NetWeight"));

                    weightQuantity = footing.GetQuantity<IfcQuantityWeight>("Qto_FootingBaseQuantities", "GrossWeight");

                    if (weightQuantity == null)
                    {
                        failedMaterials.Add((footing.EntityLabel.ToString(), "Qto_FootingBaseQuantities-GrossWeight"));

                        if (density != 0)
                        {
                            (double, string) volumeQuantity = GetFootingQuantity(footing, "Volume", density);

                            if (volumeQuantity.Item2 == "N/A")
                            {
                                return (0, "N/A");
                            }

                            UnitConverter unitConverter = new UnitConverter(volumeQuantity.Item2, "m3");

                            double volume = unitConverter.Convert(volumeQuantity.Item1);

                            weight = volume * density;

                            return (weight, "kg");
                        }

                        return (0, "N/A");
                    }
                }

                weight = weightQuantity.WeightValue;

                string unit = "kg";

                if (weightQuantity.Unit != null)
                {
                    unit = weightQuantity.Unit.ToString();
                }

                return (weight, unit);
            }
            else if (unitCategory == ".ea")
            {
                return (1, ".ea");
            }

            return (0, "N/A");
        }//L

        private (double, string) GetColumnQuantity(IfcColumn column, string unitCategory, double density)
        {
            if (unitCategory == "Length")
            {
                var lengthQuantity = column.GetQuantity<IfcQuantityLength>("Qto_ColumnBaseQuantities", "Length");

                if (lengthQuantity == null)
                {
                    failedMaterials.Add((column.EntityLabel.ToString(), "Qto_ColumnBaseQuantities-Length"));

                    return (0, "N/A");
                }

                return (lengthQuantity.LengthValue, lengthQuantity.Unit.Name());
            }
            else if (unitCategory == "Volume")
            {
                var volumeQuantity = column.GetQuantity<IfcQuantityVolume>("Qto_ColumnBaseQuantities", "NetVolume");

                if (volumeQuantity == null)
                {
                    failedMaterials.Add((column.EntityLabel.ToString(), "Qto_ColumnBaseQuantities-NetVolume"));

                    volumeQuantity = column.GetQuantity<IfcQuantityVolume>("Qto_ColumnBaseQuantities", "GrossVolume");

                    if (volumeQuantity == null)
                    {
                        failedMaterials.Add((column.EntityLabel.ToString(), "Qto_ColumnBaseQuantities-GrossVolume"));

                        var lengthQuantity = column.GetQuantity<IfcQuantityLength>("Qto_ColumnBaseQuantities", "Length");

                        var areaQuantity = column.GetQuantity<IfcQuantityArea>("Qto_ColumnBaseQuantities", "CrossSectionArea");

                        if (lengthQuantity != null && areaQuantity != null)
                        {
                            UnitConverter unitConverter = new UnitConverter(lengthQuantity.Unit.Name(), "m");

                            double length = unitConverter.Convert(lengthQuantity.LengthValue);

                            unitConverter = new UnitConverter(areaQuantity.Unit.Name(), "m2");

                            double area = unitConverter.Convert(areaQuantity.AreaValue);

                            return (length * area, "m3");
                        }
                        else
                        {
                            failedMaterials.Add((column.EntityLabel.ToString(), "Qto_ColumnBaseQuantities-Length"));

                            failedMaterials.Add((column.EntityLabel.ToString(), "Qto_ColumnBaseQuantities-CrossSectionArea"));

                            return (0, "N/A");
                        }
                    }
                }

                double volume = volumeQuantity.VolumeValue;

                string unit = "m3";

                if (volumeQuantity.Unit != null)
                {
                    unit = volumeQuantity.Unit.ToString();
                }

                return (volume, unit);
            }
            else if (unitCategory == "Weight")
            {
                double weight = 0;

                var weightQuantity = column.GetQuantity<IfcQuantityWeight>("Qto_ColumnBaseQuantities", "NetWeight");

                if (weightQuantity == null)
                {
                    failedMaterials.Add((column.EntityLabel.ToString(), "Qto_ColumnBaseQuantities-NetWeight"));

                    weightQuantity = column.GetQuantity<IfcQuantityWeight>("Qto_ColumnBaseQuantities", "GrossWeight");

                    if (weightQuantity == null)
                    {
                        failedMaterials.Add((column.EntityLabel.ToString(), "Qto_ColumnBaseQuantities-GrossWeight"));

                        if (density != 0)
                        {
                            (double, string) volumeQuantity = GetColumnQuantity(column, "Volume", density);

                            if (volumeQuantity.Item2 == "N/A")
                            {
                                return (0, "N/A");
                            }

                            UnitConverter unitConverter = new UnitConverter(volumeQuantity.Item2, "m3");

                            double volume = unitConverter.Convert(volumeQuantity.Item1);

                            weight = volume * density;

                            return (weight, "kg");
                        }

                        return (0, "N/A");
                    }
                }

                weight = weightQuantity.WeightValue;

                string unit = "kg";

                if (weightQuantity.Unit != null)
                {
                    unit = weightQuantity.Unit.ToString();
                }

                return (weight, unit);
            }
            else if (unitCategory == ".ea")
            {
                return (1, ".ea");
            }

            return (0, "N/A");
        }//L

        private (double, string) GetBeamQuantity(IfcBeam beam, string unitCategory, double density)
        {
            if (unitCategory == "Length")
            {
                var lengthQuantity = beam.GetQuantity<IfcQuantityLength>("Qto_BeamBaseQuantities", "Length");

                if (lengthQuantity == null)
                {
                    failedMaterials.Add((beam.EntityLabel.ToString(), "Qto_BeamBaseQuantities-Length"));

                    return (0, "N/A");
                }

                string unit = "mm";

                if (lengthQuantity.Unit != null)
                {
                    unit = lengthQuantity.Unit.ToString();
                }

                return (lengthQuantity.LengthValue, unit);
            }
            else if (unitCategory == "Volume")
            {
                var volumeQuantity = beam.GetQuantity<IfcQuantityVolume>("Qto_BeamBaseQuantities", "NetVolume");

                if (volumeQuantity == null)
                {
                    failedMaterials.Add((beam.EntityLabel.ToString(), "Qto_BeamBaseQuantities-NetVolume"));

                    volumeQuantity = beam.GetQuantity<IfcQuantityVolume>("Qto_BeamBaseQuantities", "GrossVolume");

                    if (volumeQuantity == null)
                    {
                        failedMaterials.Add((beam.EntityLabel.ToString(), "Qto_BeamBaseQuantities-GrossVolume"));

                        var lengthQuantity = beam.GetQuantity<IfcQuantityLength>("Qto_BeamBaseQuantities", "Length");

                        var areaQuantity = beam.GetQuantity<IfcQuantityArea>("Qto_BeamBaseQuantities", "CrossSectionArea");

                        if (lengthQuantity != null && areaQuantity != null)
                        {
                            UnitConverter unitConverter = new UnitConverter(lengthQuantity.Unit.Name(), "m");

                            double length = unitConverter.Convert(lengthQuantity.LengthValue);

                            unitConverter = new UnitConverter(areaQuantity.Unit.Name(), "m2");

                            double area = unitConverter.Convert(areaQuantity.AreaValue);

                            return (length * area, "m3");
                        }
                        else
                        {
                            failedMaterials.Add((beam.EntityLabel.ToString(), "Qto_BeamBaseQuantities-Length"));

                            failedMaterials.Add((beam.EntityLabel.ToString(), "Qto_BeamBaseQuantities-CrossSectionArea"));

                            return (0, "N/A");
                        }
                    }
                }

                double volume = volumeQuantity.VolumeValue;

                string unit = "m3";

                if (volumeQuantity.Unit != null)
                {
                    unit = volumeQuantity.Unit.ToString();
                }

                return (volume, unit);
            }
            else if (unitCategory == "Weight")
            {
                double weight = 0;

                var weightQuantity = beam.GetQuantity<IfcQuantityWeight>("Qto_BeamBaseQuantities", "NetWeight");

                if (weightQuantity == null)
                {
                    failedMaterials.Add((beam.EntityLabel.ToString(), "Qto_BeamBaseQuantities-NetWeight"));

                    weightQuantity = beam.GetQuantity<IfcQuantityWeight>("Qto_BeamBaseQuantities", "GrossWeight");

                    if (weightQuantity == null)
                    {
                        failedMaterials.Add((beam.EntityLabel.ToString(), "Qto_BeamBaseQuantities-GrossWeight"));

                        if (density != 0)
                        {
                            (double, string) volumeQuantity = GetBeamQuantity(beam, "Volume", density);

                            if (volumeQuantity.Item2 == "N/A")
                            {
                                return (0, "N/A");
                            }

                            UnitConverter unitConverter = new UnitConverter(volumeQuantity.Item2, "m3");

                            double volume = unitConverter.Convert(volumeQuantity.Item1);

                            weight = volume * density;

                            return (weight, "kg");
                        }

                        return (0, "N/A");
                    }
                }

                weight = weightQuantity.WeightValue;

                string unit = "kg";

                if (weightQuantity.Unit != null)
                {
                    unit = weightQuantity.Unit.ToString();
                }

                return (weight, unit);
            }
            else if (unitCategory == ".ea")
            {
                return (1, ".ea");
            }

            return (0, "N/A");
        }//L

        private (double, string) GetRampFlightQuantity(IfcRampFlight rampFlight, string unitCategory, double density)
        {
            if (unitCategory == "Length")
            {
                var lengthQuantity = rampFlight.GetQuantity<IfcQuantityLength>("Qto_RampFlightBaseQuantities", "Length");

                if (lengthQuantity == null)
                {
                    failedMaterials.Add((rampFlight.EntityLabel.ToString(), "Qto_RampFlightBaseQuantities-Length"));

                    return (0, "N/A");
                }

                string unit = "mm";

                if (lengthQuantity.Unit != null)
                {
                    unit = lengthQuantity.Unit.ToString();
                }

                return (lengthQuantity.LengthValue, unit);
            }
            else if (unitCategory == "Area")
            {
                var areaQuantity = rampFlight.GetQuantity<IfcQuantityArea>("Qto_RampFlightBaseQuantities", "NetArea");

                if (areaQuantity == null)
                {
                    failedMaterials.Add((rampFlight.EntityLabel.ToString(), "Qto_RampFlightBaseQuantities-NetArea"));

                    areaQuantity = rampFlight.GetQuantity<IfcQuantityArea>("Qto_RampFlightBaseQuantities", "GrossArea");

                    if (areaQuantity == null)
                    {
                        failedMaterials.Add((rampFlight.EntityLabel.ToString(), "Qto_RampFlightBaseQuantities-GrossArea"));

                        var lengthQuantity = rampFlight.GetQuantity<IfcQuantityLength>("Qto_RampFlightBaseQuantities", "Length");

                        var widthQuantity = rampFlight.GetQuantity<IfcQuantityLength>("Qto_RampFlightBaseQuantities", "Width");

                        if (widthQuantity != null && lengthQuantity != null)
                        {
                            UnitConverter unitConverter = new UnitConverter(lengthQuantity.Unit.Name(), "m");

                            double length = unitConverter.Convert(lengthQuantity.LengthValue);

                            unitConverter = new UnitConverter(widthQuantity.Unit.Name(), "m");

                            double width = unitConverter.Convert(widthQuantity.LengthValue);

                            return (length * width, "m2");
                        }
                        else
                        {
                            failedMaterials.Add((rampFlight.EntityLabel.ToString(), "Qto_RampFlightBaseQuantities-Length"));

                            failedMaterials.Add((rampFlight.EntityLabel.ToString(), "Qto_RampFlightBaseQuantities-Length"));

                            return (0, "N/A");
                        }
                    }
                }

                double area = areaQuantity.AreaValue;

                string unit = "m2";

                if (areaQuantity.Unit != null)
                {
                    unit = areaQuantity.Unit.ToString();
                }

                return (area, unit);
            }
            else if (unitCategory == "Volume")
            {
                var volumeQuantity = rampFlight.GetQuantity<IfcQuantityVolume>("Qto_StairFlightBaseQuantities", "NetVolume");

                double volume = 0;

                if (volumeQuantity == null)
                {
                    failedMaterials.Add((rampFlight.EntityLabel.ToString(), "Qto_StairFlightBaseQuantities-NetVolume"));

                    volumeQuantity = rampFlight.GetQuantity<IfcQuantityVolume>("Qto_StairFlightBaseQuantities", "GrossVolume");

                    if (volumeQuantity == null)
                    {
                        failedMaterials.Add((rampFlight.EntityLabel.ToString(), "Qto_StairFlightBaseQuantities-GrossVolume"));

                        var depth = rampFlight.GetQuantity<IfcQuantityLength>("Qto_RampFlightBaseQuantities", "Width");

                        var areaQuantity = GetRampFlightQuantity(rampFlight, "Area", density);

                        if (areaQuantity.Item2 == "N/A")
                        {
                            return (0, "N/A");
                        }

                        UnitConverter unitConverter_1 = new UnitConverter(areaQuantity.Item2, "m2");

                        double area = unitConverter_1.Convert(areaQuantity.Item1);

                        UnitConverter unitConverter_2 = new UnitConverter(depth.Unit.Name(), "m");

                        double thickness = unitConverter_2.Convert(depth.LengthValue);

                        volume = area * thickness;

                        return (volume, "m3");
                    }
                }

                volume = volumeQuantity.VolumeValue;

                string unit = "m3";

                if (volumeQuantity.Unit != null)
                {
                    unit = volumeQuantity.Unit.ToString();
                }

                return (volume, unit);
            }
            else if (unitCategory == "Weight")
            {
                double weight = 0;

                if (density != 0)
                {
                    (double, string) volumeQuantity = GetRampFlightQuantity(rampFlight, "Volume", density);

                    if (volumeQuantity.Item2 == "N/A")
                    {
                        return (0, "N/A");
                    }

                    UnitConverter unitConverter = new UnitConverter(volumeQuantity.Item2, "m3");

                    double volume = unitConverter.Convert(volumeQuantity.Item1);

                    weight = volume * density;

                    return (weight, "kg");
                }
            }
            else if (unitCategory == ".ea")
            {
                return (1, ".ea");
            }

            return (0, "N/A");
        }//L

        private (double, string) GetStairFlightQuantity(IfcStairFlight stairFlight, string unitCategory, double density)
        {

            if (unitCategory == "Length")
            {
                var lengthQuantity = stairFlight.GetQuantity<IfcQuantityLength>("Qto_StairFlightBaseQuantities", "Length");

                if (lengthQuantity == null)
                {
                    failedMaterials.Add((stairFlight.EntityLabel.ToString(), "Qto_StairFlightBaseQuantities-Length"));

                    return (0, "N/A");
                }

                string unit = "mm";

                if (lengthQuantity.Unit != null)
                {
                    unit = lengthQuantity.Unit.ToString();
                }

                return (lengthQuantity.LengthValue, unit);
            }
            else if (unitCategory == "Volume")
            {
                var volumeQuantity = stairFlight.GetQuantity<IfcQuantityVolume>("Qto_StairFlightBaseQuantities", "NetVolume");

                if (volumeQuantity == null)
                {
                    failedMaterials.Add((stairFlight.EntityLabel.ToString(), "Qto_StairFlightBaseQuantities-NetVolume"));

                    volumeQuantity = stairFlight.GetQuantity<IfcQuantityVolume>("Qto_StairFlightBaseQuantities", "GrossVolume");

                    if (volumeQuantity == null)
                    {
                        failedMaterials.Add((stairFlight.EntityLabel.ToString(), "Qto_StairFlightBaseQuantities-GrossVolume"));

                        return (0, "N/A");
                    }
                }

                double volume = volumeQuantity.VolumeValue;

                string unit = "m3";

                if (volumeQuantity.Unit != null)
                {
                    unit = volumeQuantity.Unit.ToString();
                }

                return (volume, unit);
            }
            else if (unitCategory == "Weight")
            {
                double weight = 0;

                if (density != 0)
                {
                    (double, string) volumeQuantity = GetStairFlightQuantity(stairFlight, "Volume", density);

                    if (volumeQuantity.Item2 == "N/A")
                    {
                        return (0, "N/A");
                    }

                    UnitConverter unitConverter = new UnitConverter(volumeQuantity.Item2, "m3");

                    double volume = unitConverter.Convert(volumeQuantity.Item1);

                    weight = volume * density;

                    return (weight, "kg");
                }
            }
            else if (unitCategory == ".ea")
            {
                return (1, ".ea");
            }

            return (0, "N/A");
        }//L

        private (double, string) GetCoveringQuantity(IfcCovering covering, string unitCategory, double thickness, double density)
        {
            if (unitCategory == "Area")
            {
                var areaQuantity = covering.GetQuantity<IfcQuantityArea>("Qto_CoveringBaseQuantities", "NetArea");

                if (areaQuantity == null)
                {
                    failedMaterials.Add((covering.EntityLabel.ToString(), "Qto_CoveringBaseQuantities-NetArea"));

                    areaQuantity = covering.GetQuantity<IfcQuantityArea>("Qto_CoveringBaseQuantities", "GrossArea");

                    if (areaQuantity == null)
                    {
                        failedMaterials.Add((covering.EntityLabel.ToString(), "Qto_CoveringBaseQuantities-GrossArea"));

                        return (0, "N/A");
                    }

                }

                double area = areaQuantity.AreaValue;

                string unit = "m2";

                if (areaQuantity.Unit != null)
                {
                    unit = areaQuantity.Unit.ToString();
                }

                return (area, unit);
            }
            else if (unitCategory == "Volume")
            {
                var areaQuantity = GetCoveringQuantity(covering, "Area", thickness, density);

                if (areaQuantity.Item2 == "N/A")
                {
                    return (0, "N/A");
                }

                var depth = covering.GetQuantity<IfcQuantityLength>("Qto_CoveringBaseQuantities", "Width");

                UnitConverter unitConverter_1 = new UnitConverter(areaQuantity.Item2, "m2");

                double area = unitConverter_1.Convert(areaQuantity.Item1);

                double volume = 0;

                if (depth == null || !CheckIfSingleLayerCover(covering))
                {
                    volume = area * thickness / 1000;
                }
                else
                {
                    UnitConverter unitConverter_2 = new UnitConverter(depth.Unit.Name(), "mm");

                    double t = unitConverter_2.Convert(depth.LengthValue);

                    volume = area * t / 1000;
                }

                return (volume, "m3");
            }
            else if (unitCategory == "Weight")
            {
                double weight = 0;

                if (density != 0)
                {
                    (double, string) volumeQuantity = GetCoveringQuantity(covering, "Volume", thickness, density);

                    if (volumeQuantity.Item2 == "N/A")
                    {
                        return (0, "N/A");
                    }

                    UnitConverter unitConverter = new UnitConverter(volumeQuantity.Item2, "m3");

                    double volume = unitConverter.Convert(volumeQuantity.Item1);

                    weight = volume * density;

                    return (weight, "kg");
                }

                return (0, "N/A");
            }
            else if (unitCategory == ".ea")
            {
                return (1, ".ea");
            }

            return (0, "N/A");
        }//L

        private bool CheckIfSingleLayerCover(IfcCovering covering)
        {
            var materialLayerSet = covering.HasAssociations.OfType<IfcRelAssociatesMaterial>().FirstOrDefault();

            if (materialLayerSet == null)
            {
                return false;
            }

            var materialLayerSetUsage = materialLayerSet.RelatingMaterial as IfcMaterialLayerSetUsage;

            if (materialLayerSetUsage == null)
            {
                return false;
            }

            var materialLayerSetDef = materialLayerSetUsage.ForLayerSet as IfcMaterialLayerSet;

            if (materialLayerSetDef == null)
            {
                return false;
            }

            return materialLayerSetDef.MaterialLayers.Count == 1;
        }

        private (double, string) GetRoofQuantity(IfcRoof roof, string unitCategory, double thickness, double density)
        {
            if (unitCategory == "Area")
            {
                var areaQuantity = roof.GetQuantity<IfcQuantityArea>("Qto_RoofBaseQuantities", "NetArea");

                if (areaQuantity == null)
                {
                    failedMaterials.Add((roof.EntityLabel.ToString(), "Qto_RoofBaseQuantities-NetArea"));

                    areaQuantity = roof.GetQuantity<IfcQuantityArea>("Qto_RoofBaseQuantities", "GrossArea");

                    if (areaQuantity == null)
                    {
                        failedMaterials.Add((roof.EntityLabel.ToString(), "Qto_RoofBaseQuantities-GrossArea"));

                        areaQuantity= roof.GetQuantity<IfcQuantityArea>("Qto_RoofBaseQuantities", "ProjectedArea");

                        if (areaQuantity == null)
                        {
                            failedMaterials.Add((roof.EntityLabel.ToString(), "Qto_RoofBaseQuantities-ProjectedArea"));

                            return (0, "N/A");
                        }                      
                    }
                }

                double area = areaQuantity.AreaValue;

                string unit = "m2";

                if (areaQuantity.Unit != null)
                {
                    unit = areaQuantity.Unit.ToString();
                }

                return (area, unit);
            }
            else if (unitCategory == "Volume")
            {
                var areaQuantity = GetRoofQuantity(roof, "Area", thickness, density);

                if (areaQuantity.Item2 == "N/A")
                {
                    return (0, "N/A");
                }

                UnitConverter unitConverter = new UnitConverter(areaQuantity.Item2, "m2");

                double area = unitConverter.Convert(areaQuantity.Item1);

                double volume = area * thickness / 1000;

                return (volume, "m3");
            }
            else if (unitCategory == "Weight")
            {
                double weight = 0;

                if (density != 0)
                {
                    (double, string) volumeQuantity = GetRoofQuantity(roof, "Volume", thickness, density);

                    if (volumeQuantity.Item2 == "N/A")
                    {
                        return (0, "N/A");
                    }

                    UnitConverter unitConverter = new UnitConverter(volumeQuantity.Item2, "m3");

                    double volume = unitConverter.Convert(volumeQuantity.Item1);

                    weight = volume * density;

                    return (weight, "kg");
                }
            }
            else if (unitCategory == ".ea")
            {
                return (1, ".ea");
            }

            return (0, "N/A");
        }//L

        private (double, string) GetSlabQuantity(IfcSlab slab, string unitCategory, double thickness, double density) //L
        {
            if (unitCategory == "Area")
            {
                var areaQuantity = slab.GetQuantity<IfcQuantityArea>("Qto_SlabBaseQuantities", "NetArea");

                if (areaQuantity == null)
                {
                    failedMaterials.Add((slab.EntityLabel.ToString(), "Qto_SlabBaseQuantities-NetArea"));

                    areaQuantity = slab.GetQuantity<IfcQuantityArea>("Qto_SlabBaseQuantities", "GrossArea");

                    if (areaQuantity == null)
                    {
                        failedMaterials.Add((slab.EntityLabel.ToString(), "Qto_SlabBaseQuantities-GrossArea"));

                        return (0, "N/A");
                    }
                }

                double area = areaQuantity.AreaValue;

                string unit = "m2";

                if (areaQuantity.Unit != null)
                {
                    unit = areaQuantity.Unit.ToString();
                }

                return (area, unit);
            }
            else if (unitCategory == "Volume")
            {
                var depth = slab.GetQuantity<IfcQuantityLength>("Qto_SlabBaseQuantities", "Depth");

                var volumeQuantity = slab.GetQuantity<IfcQuantityVolume>("Qto_SlabBaseQuantities", "NetVolume");

                if (volumeQuantity == null)
                {
                    volumeQuantity = slab.GetQuantity<IfcQuantityVolume>("Qto_SlabBaseQuantities", "GrossVolume");
                }

                if(CheckIfSingleLayerSlab(slab))
                {
                    if(volumeQuantity != null)
                    {
                        double volume = volumeQuantity.VolumeValue;

                        string unit = "m3";

                        if (volumeQuantity.Unit != null)
                        {
                            unit = volumeQuantity.Unit.ToString();
                        }

                        return (volume, unit);
                    }
                    else
                    {
                        var areaQuantity = GetSlabQuantity(slab, "Area", thickness, density);

                        if (areaQuantity.Item2 == "N/A")
                        {
                            return (0, "N/A");
                        }

                        UnitConverter unitConverter = new UnitConverter(areaQuantity.Item2, "m2");

                        double area = unitConverter.Convert(areaQuantity.Item1);

                        return (area * thickness / 1000, "m3");
                    }
                }
                else
                {
                    if (volumeQuantity == null || depth == null)
                    {
                        var areaQuantity = GetSlabQuantity(slab, "Area", thickness, density);

                        if (areaQuantity.Item2 == "N/A")
                        {
                            return (0, "N/A");
                        }

                        UnitConverter unitConverter = new UnitConverter(areaQuantity.Item2, "m2");

                        double area = unitConverter.Convert(areaQuantity.Item1);

                        return (area * thickness / 1000, "m3");

                    }
                    else
                    {
                        double volumeTotal = volumeQuantity.VolumeValue;

                        double volume = volumeTotal * thickness / depth.LengthValue;

                        string unit = "m3";

                        if (volumeQuantity.Unit != null)
                        {
                            unit = volumeQuantity.Unit.ToString();
                        }

                        return (volume, unit);
                    }
                }
            }
            else if (unitCategory == "Weight")
            {
                double weight = 0;

                string unit = "N/A";

                bool isSingleLayer = CheckIfSingleLayerSlab(slab);

                var depth = slab.GetQuantity<IfcQuantityLength>("Qto_SlabBaseQuantities", "Depth");

                var weightQuantity = slab.GetQuantity<IfcQuantityWeight>("Qto_SlabBaseQuantities", "NetWeight");

                if (weightQuantity == null)
                {
                    weightQuantity = slab.GetQuantity<IfcQuantityWeight>("Qto_SlabBaseQuantities", "GrossWeight");
                }

                if(isSingleLayer && weightQuantity != null)
                {
                    weight = weightQuantity.WeightValue;

                    unit = "kg";

                    if (weightQuantity.Unit != null)
                    {
                        unit = weightQuantity.Unit.ToString();
                    }

                    return (weight, unit);
                }
                else
                {
                    if (density != 0)
                    {
                        (double, string) volumeQuantity = GetSlabQuantity(slab, "Volume", thickness, density);

                        if (volumeQuantity.Item2 == "N/A")
                        {
                            return (0, "N/A");
                        }

                        UnitConverter unitConverter = new UnitConverter(volumeQuantity.Item2, "m3");

                        double volume = unitConverter.Convert(volumeQuantity.Item1);

                        weight = volume * density;

                        return (weight, "kg");
                    }

                    return (0, "N/A");
                }
            }
            else if (unitCategory == ".ea")
            {
                return (1, ".ea");
            }

            return (0, "N/A");
        }

        private bool CheckIfSingleLayerSlab(IfcSlab slab)
        {
            var materialLayerSet = slab.HasAssociations.OfType<IfcRelAssociatesMaterial>().FirstOrDefault();

            if (materialLayerSet == null)
            {
                return false;
            }

            var materialLayerSetUsage = materialLayerSet.RelatingMaterial as IfcMaterialLayerSetUsage;

            if (materialLayerSetUsage == null)
            {
                return false;
            }

            var materialLayerSetDef = materialLayerSetUsage.ForLayerSet as IfcMaterialLayerSet;

            if (materialLayerSetDef == null)
            {
                return false;
            }

            return materialLayerSetDef.MaterialLayers.Count == 1;
        }

        private (double, string) GetCurtainWallQuantity(IfcCurtainWall curtainWall, string unitCategory, double thickness, double density)
        {
            if (unitCategory == "Area")
            {
                var areaQuantity = curtainWall.GetQuantity<IfcQuantityArea>("Qto_CurtainWallQuantities", "NetSideArea");

                if (areaQuantity == null)
                {
                    failedMaterials.Add((curtainWall.EntityLabel.ToString(), "Qto_CurtainWallQuantities-NetSideArea"));

                    areaQuantity = curtainWall.GetQuantity<IfcQuantityArea>("Qto_CurtainWallQuantities", "GrossSideArea");

                    if (areaQuantity == null)
                    {
                        failedMaterials.Add((curtainWall.EntityLabel.ToString(), "Qto_CurtainWallQuantities-GrossSideArea"));

                        return (0, "N/A");
                    }
                }

                double area = areaQuantity.AreaValue;

                string unit = "m2";

                if (areaQuantity.Unit != null)
                {
                    unit = areaQuantity.Unit.ToString();
                }

                return (area, unit);
            }
            else if (unitCategory == "Length")
            {
                var lengthQuantity = curtainWall.GetQuantity<IfcQuantityLength>("Qto_CurtainWallQuantities", "Length");

                if (lengthQuantity == null)
                {
                    failedMaterials.Add((curtainWall.EntityLabel.ToString(), "Qto_CurtainWallQuantities-Length"));

                    return (0, "N/A");
                }

                double length = lengthQuantity.LengthValue;

                string unit = "mm";

                if (lengthQuantity.Unit != null)
                {
                    unit = lengthQuantity.Unit.ToString();
                }

                return (length, unit);
            }
            else if (unitCategory == ".ea")
            {
                return (1, ".ea");
            }
            else if (unitCategory=="Weight")
            {
                var areaQuantity = GetCurtainWallQuantity(curtainWall, "Area", thickness, density);

                if (areaQuantity.Item2 == "N/A")
                {
                    return (0, "N/A");
                }

                UnitConverter unitConverter = new UnitConverter(areaQuantity.Item2, "m2");

                double area = unitConverter.Convert(areaQuantity.Item1);

                double weight = 0;

                if (density != 0)
                {
                    weight = area * thickness / 1000 * density;
                }

                return (weight, "kg");

            }

            return (0, "N/A");
        }//L

        private (double, string) GetWallQuantity(IfcWall wall, string unitCategory, double thickness, double density)
        {
            if (unitCategory == "Area")
            {
                var areaQuantity = wall.GetQuantity<IfcQuantityArea>("Qto_WallBaseQuantities", "NetSideArea");

                if (areaQuantity == null)
                {
                    failedMaterials.Add((wall.EntityLabel.ToString(), "Qto_WallBaseQuantities-NetSideArea"));

                    areaQuantity = wall.GetQuantity<IfcQuantityArea>("Qto_WallBaseQuantities", "GrossSideArea");

                    if (areaQuantity == null)
                    {
                        failedMaterials.Add((wall.EntityLabel.ToString(), "Qto_WallBaseQuantities-GrossSideArea"));

                        var volumeQuantity = wall.GetQuantity<IfcQuantityVolume>("Qto_WallBaseQuantities", "NetVolume");

                        var depth = wall.GetQuantity<IfcQuantityLength>("Qto_WallBaseQuantities", "Width");

                        if(volumeQuantity != null && depth != null)
                        {
                            double volume = volumeQuantity.VolumeValue;

                            double width = depth.LengthValue / 1000;

                            return ((double)volume / width, "m2");
                        }

                        return (0, "N/A");
                    }
                }

                double area = areaQuantity.AreaValue;

                string unit = "m2";

                if (areaQuantity.Unit != null)
                {
                    unit = areaQuantity.Unit.ToString();
                }

                return (area, unit);
            }
            else if (unitCategory == "Volume")
            {
                var depth = wall.GetQuantity<IfcQuantityLength>("Qto_WallBaseQuantities", "Width");

                var volumeQuantity = wall.GetQuantity<IfcQuantityVolume>("Qto_WallBaseQuantities", "NetVolume");

                if (volumeQuantity == null)
                {
                    volumeQuantity = wall.GetQuantity<IfcQuantityVolume>("Qto_WallBaseQuantities", "GrossVolume");
                }

                if (CheckIfSingleLayerWall(wall))
                {
                    if (volumeQuantity != null)
                    {
                        double volume = volumeQuantity.VolumeValue;

                        string unit = "m3";

                        if (volumeQuantity.Unit != null)
                        {
                            unit = volumeQuantity.Unit.ToString();
                        }

                        return (volume, unit);
                    }
                    else
                    {
                        var areaQuantity = GetWallQuantity(wall, "Area", thickness, density);

                        if (areaQuantity.Item2 == "N/A")
                        {
                            return (0, "N/A");
                        }

                        UnitConverter unitConverter = new UnitConverter(areaQuantity.Item2, "m2");

                        double area = unitConverter.Convert(areaQuantity.Item1);

                        return (area * thickness / 1000, "m3");
                    }
                }
                else
                {
                    if (volumeQuantity == null || depth == null)
                    {
                        var areaQuantity = GetWallQuantity(wall, "Area", thickness, density);

                        if (areaQuantity.Item2 == "N/A")
                        {
                            return (0, "N/A");
                        }

                        UnitConverter unitConverter = new UnitConverter(areaQuantity.Item2, "m2");

                        double area = unitConverter.Convert(areaQuantity.Item1);

                        return (area * thickness / 1000, "m3");

                    }
                    else
                    {
                        double volumeTotal = volumeQuantity.VolumeValue;

                        double volume = volumeTotal * thickness / depth.LengthValue;

                        string unit = "m3";

                        if (volumeQuantity.Unit != null)
                        {
                            unit = volumeQuantity.Unit.ToString();
                        }

                        return (volume, unit);
                    }
                }
            }
            else if (unitCategory == "Length")
            {
                var lengthQuantity = wall.GetQuantity<IfcQuantityLength>("Qto_WallBaseQuantities", "Length");

                if (lengthQuantity == null)
                {
                    failedMaterials.Add((wall.EntityLabel.ToString(), "Qto_WallBaseQuantities-Length"));

                    return (0, "N/A");
                }

                double length = lengthQuantity.LengthValue;

                string unit = "mm";

                if (lengthQuantity.Unit != null)
                {
                    unit = lengthQuantity.Unit.ToString();
                }

                return (length, unit);
            }
            else if (unitCategory == "Weight")
            {
                double weight = 0;

                string unit = "N/A";

                bool isSingleLayer = CheckIfSingleLayerWall(wall);

                var depth = wall.GetQuantity<IfcQuantityLength>("Qto_WallBaseQuantities", "Width");

                var weightQuantity = wall.GetQuantity<IfcQuantityWeight>("Qto_WallBaseQuantities", "NetWeight");

                if (weightQuantity == null)
                {
                    weightQuantity = wall.GetQuantity<IfcQuantityWeight>("Qto_WallBaseQuantities", "GrossWeight");
                }

                if (isSingleLayer && weightQuantity != null)
                {
                    weight = weightQuantity.WeightValue;

                    unit = "kg";

                    if (weightQuantity.Unit != null)
                    {
                        unit = weightQuantity.Unit.ToString();
                    }

                    return (weight, unit);
                }
                else
                {
                    if (density != 0)
                    {
                        (double, string) volumeQuantity = GetWallQuantity(wall, "Volume", thickness, density);

                        if (volumeQuantity.Item2 == "N/A")
                        {
                            return (0, "N/A");
                        }

                        UnitConverter unitConverter = new UnitConverter(volumeQuantity.Item2, "m3");

                        double volume = unitConverter.Convert(volumeQuantity.Item1);

                        weight = volume * density;

                        return (weight, "kg");
                    }

                    return (0, "N/A");
                }
            }
            else if (unitCategory == ".ea")
            {
                return (1, ".ea");
            }

            return (0, "N/A");
        }

        private bool CheckIfSingleLayerWall(IfcWall wall)
        {
            var materialLayerSet = wall.HasAssociations.OfType<IfcRelAssociatesMaterial>().FirstOrDefault();

            if (materialLayerSet == null)
            {
                return false;
            }

            var materialLayerSetUsage = materialLayerSet.RelatingMaterial as IfcMaterialLayerSetUsage;

            if (materialLayerSetUsage == null)
            {
                return false;
            }

            var materialLayerSetDef = materialLayerSetUsage.ForLayerSet as IfcMaterialLayerSet;

            if (materialLayerSetDef == null)
            {
                return false;
            }

            return materialLayerSetDef.MaterialLayers.Count == 1;
        }

        private (double, string) GetWindowQuantity(IfcWindow window, string unitCategory)
        {
            if (unitCategory == "Area")
            {
                var areaQuantity = window.GetQuantity<IfcQuantityArea>("Qto_WindowBaseQuantities", "Area");

                if (areaQuantity == null)
                {
                    failedMaterials.Add((window.EntityLabel.ToString(), "Qto_WindowBaseQuantities-Area"));

                    return (0, "N/A");
                }

                double area = areaQuantity.AreaValue;

                string unit = "m2";

                if (areaQuantity.Unit != null)
                {
                    unit = areaQuantity.Unit.ToString();
                }

                return (area, unit);
            }
            else if (unitCategory == ".ea")
            {
                return (1, ".ea");
            }

            return (0, "N/A");
        }//L

        private (double, string) GetDoorQuantity(IfcDoor door, string unitCategory)
        {
            if (unitCategory == "Area")
            {
                var areaQuantity = door.GetQuantity<IfcQuantityArea>("Qto_DoorBaseQuantities", "Area");
                //         ElementQuantities.ToList().Find(x => x.Name == "Area") as IfcQuantityArea;

                if (areaQuantity == null)
                {
                    failedMaterials.Add((door.EntityLabel.ToString(), "Qto_DoorBaseQuantities-Area"));

                    return (0, "N/A");
                }

                double area = areaQuantity.AreaValue;

                string unit = "m2";

                if(areaQuantity.Unit != null)
                {
                    unit=areaQuantity.Unit.ToString();
                }
              
                /*var quantitySet=door.IsDefinedBy.ToList().Find(x=>x.RelatingPropertyDefinition.GetType().Name == "IfcElementQuantity").RelatingPropertyDefinition as IfcElementQuantity;

                if (quantitySet == null)
                {
                    MessageBox.Show("Coundn't found the Quantity Set!");

                    return (0, "N/A");
                }

                var Area = quantitySet.Quantities.ToList().Find(x => x.Name == "Area") as IfcQuantityArea;

                if(Area == null)
                {
                    MessageBox.Show("Coundn't found the Area!");

                    return (0, "N/A");
                }

                double area_2 = Area.AreaValue;*/

                return (area, unit);
            }
            else if (unitCategory == ".ea")
            {
                return (1, ".ea");
            }

            return (0, "N/A");
        }//L

        private double RetriveWeight(string uid, string IfcType)
        {
            string sql = $"SELECT Density FROM [dbo].[{IfcType} Table] WHERE Material_ID='{uid}';";

            SqlCommand query = connect.Query(sql);

            decimal weight = 0;

            using (SqlDataReader reader = query.ExecuteReader())
            {
                if (reader.Read())
                {
                    weight = reader.GetDecimal(0);
                }
            }

            return (double)weight;
        }

        private string DetermineUnitCategory(string Unit)
        {
            switch (Unit.ToLower().Trim())
            {
                case "m3":
                case "m^3":
                case "m³":
                case "cubic meter":
                case "cubic meters":
                case "cubic metre":
                case "cubic metres":
                case "cm3":
                case "cm^3":
                case "cm³":
                case "cubic centimeter":
                case "cubic centimeters":
                case "cubic centimetre":
                case "cubic centimetres":
                case "l":
                case "liter":
                case "liters":
                case "litre":
                case "litres":
                case "gal":
                case "gallon":
                case "gallons":
                case "ft3":
                case "ft^3":
                case "ft³":
                case "cubic foot":
                case "cubic feet":
                case "yd3":
                case "yd^3":
                case "yd³":
                case "cubic yard":
                case "cubic yards":
                case "in3":
                case "in^3":
                case "in³":
                case "cubic inch":
                case "cubic inches":
                    return "Volume";
                case "kg":
                case "kilogram":
                case "kilograms":
                case "g":
                case "gram":
                case "grams":
                case "t":
                case "ton":
                case "tons":
                case "tonne":
                case "tonnes":
                case "lb":
                case "lbs":
                case "pound":
                case "pounds":
                case "oz":
                case "ounce":
                case "ounces":
                    return "Weight";
                case "m2":
                case "m^2":
                case "m²":
                case "square meter":
                case "square meters":
                case "square metre":
                case "square metres":
                case "cm2":
                case "cm^2":
                case "cm²":
                case "square centimeter":
                case "square centimeters":
                case "square centimetre":
                case "square centimetres":
                    return "Area";
                case "m":
                case "meter":
                case "meters":
                case "metre":
                case "metres":
                case "cm":
                case "centimeter":
                case "centimeters":
                case "centimetre":
                case "centimetres":
                case "mm":
                case "millimeter":
                case "millimeters":
                case "millimetre":
                case "millimetres":
                case "ft":
                case "foot":
                case "feet":
                case "yd":
                case "yard":
                case "yards":
                case "in":
                case "inch":
                case "inches":
                case "km":
                case "kms":
                case "kilometer":
                case "kilometers":
                case "kilometre":
                case "kilometres":
                case "mile":
                case "miles":
                case "mi":
                case "nautical mile":
                case "nautical miles":
                    return "Length";
                default:
                    return Unit;
            }
        }

        #endregion

        #endregion
      
    }

    #region DataStrcutures
    public class Record
    {
        public string Uid { get; set; }
        public string Unit { get; set; }
        public string MaterialType { get; set; }
        public Double GWP { get; set; }
        public Double ODP { get; set; }
        public Double AP { get; set; }
        public Double EP { get; set; }
        public Double POCP { get; set; }
        public Double ADPE { get; set; }
        public Double ADPF { get; set; }
        public Double PERT { get; set; }
        public Double PENRT { get; set; }
        public string Address { get; set; }
    }

    public class MaterialInfo
    {
        public string IfcLabel { get; set; }

        public string Name { get; set; }

        public string IfcType { get; set; }

        public string ElementType { get; set; }

        public string ID { get; set; }

        public List<Record> MaterialRecords { get; set; }

        public double quantity { get; set; }

        public string unit { get; set; }
    }

    public class EleLCAResult
    {
        public string IfcLabel { get; set; }

        public string Name { get; set; }

        public string Stage { get; set; }

        public string Description { get; set; }

        public int materialCount { get; set; }

        public Dictionary<string, List<Record>> MaterialResults { get; set; }

        public Dictionary<string, Record> AverageResults { get; set; }

        public Dictionary<string, Record> MinResults { get; set; }

        public Dictionary<string, Record> MaxResults { get; set; }

        public Dictionary<string, Record> StandardDeviation { get; set; }

       // public List<Record> TotalResults { get; set; }

        public Record TotalResult_SD { get; set; }

        public Record TotalResult_Min { get; set; }

        public Record TotalResult_Max { get; set; }

        public Record TotalResult_Average { get; set; }
    }

    #endregion
}
