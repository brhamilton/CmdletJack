using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using Microsoft.PowerShell.Commands;

namespace CmdletJack
{
    class Program
    {
        private static XNamespace nsMML = "http://schemas.microsoft.com/maml/2004/10";
        private static XNamespace nsCMD = "http://schemas.microsoft.com/maml/dev/command/2004/10";
        private static XNamespace nsDEV = "http://schemas.microsoft.com/maml/dev/2004/10";
        private static XDocument xd;
        private static List<CmdletData> ProjectCmdlets = new List<CmdletData>();
        private static List<string> CommonParams = new List<string>();
        private static List<string> ValidReportTypes = new List<string>();
        private static List<string> ApprovedCompanyNames = new List<string>();
        private static List<string> JobList;
        private static List<string> headersCmdletParamData = new List<string>();
        private static List<string> headersCmdletDescrData = new List<string>();
        private static List<string> headersCmdletExampleData = new List<string>();
        private static List<string> headersCmdletInOutData = new List<string>();
        private static List<string> headersCmdletNameHitsData = new List<string>();
        private static List<string> headersCmdletNameHitsContextData = new List<string>();
        private static List<string> headersCmdletLinkData = new List<string>();
        private static List<string> headersSearchResults = new List<string>();
        private static List<string> headersCmdletMetrics = new List<string>();
        private static List<string> headersSummary = new List<string>();
        private static List<string> headersWarning = new List<string>();
        private static Dictionary<string, string> HeaderRow = new Dictionary<string, string>();

        private static string projName = String.Empty;
        private static string saveDir = String.Empty;
        private static FileInfo fi;

        private static string searchPattern = String.Empty;
        private static bool matchcase = false;
        private static bool wholeword = false;
        private static bool append = false;
        private static int curSearchHits = 0;

        private static int ProjParamCount;
        private static bool allProjsFinished = false;

        private static CmdletMetrics Tallies;
        private static List<CmdletMetrics> TalliesList = new List<CmdletMetrics>();
        private static List<Summary> SummaryList = new List<Summary>();
        private static List<CmdletDescrData> DescrList = new List<CmdletDescrData>();
        private static List<CmdletParamData> ParamsList = new List<CmdletParamData>();
        private static List<CmdletExampleData> ExampleList = new List<CmdletExampleData>();
        private static List<CmdletInputOutputData> InOutList = new List<CmdletInputOutputData>();
        private static List<CmdletNameHitsData> NameHitList = new List<CmdletNameHitsData>();
        private static List<CmdletLinkData> LinksList = new List<CmdletLinkData>();
        private static List<SearchResults> SearchResultList = new List<SearchResults>();

        private static List<Warning> Warns = new List<Warning>();

        private static string arg1 = string.Empty;
        private static string arg2 = string.Empty;
        private static string arg3 = string.Empty;
        private static string arg4 = string.Empty;

        #region Processing and reporting methods

        private static void Main(string[] args)
        {
            try
            {
                FillValidReportTypes();
                FillCommonParameters();
                FillApprovedCompanyNames();
                FillHeaders();

                if (args.Length == 0)
                {
                    ShowUsage();

                }
                else
                {
                    foreach (string a in args)
                    {
                        switch (args.Length)
                        {
                            case 0:

                                break;
                            case 1:
                                arg1 = args[0];
                                break;
                            case 2:
                                arg1 = args[0];
                                arg2 = args[1];
                                break;
                            case 3:
                                arg1 = args[0];
                                arg2 = args[1];
                                arg3 = args[2];
                                break;
                            case 4:
                                arg1 = args[0];
                                arg2 = args[1];
                                arg3 = args[2];
                                arg4 = args[3];
                                break;
                            default:
                                break;
                        }
                    }
                    if (args.Length == 1)
                    {
                        // Args layout -- L = length
                        // L = 0      L = 1             
                        // CmdletJack <C:\Path\Cmdlets>
                        //            arg1

                        // user asked for usage, specified a path/file for cmdlets without -a,
                        // incompleted a search, or some other error

                        // Remember, collection is zero based but arg variables start at 1

                        if (arg1 == "?" || arg1 == "/?" || arg1 == "-?" || arg1 == @"\?")
                        {
                            ShowUsage();
                        }
                        else if (Directory.Exists(arg1) || File.Exists(arg1))
                        {
                            StartJob(arg1);
                        }
                        else
                        {
                            Console.WriteLine("Invalid directory or file specification.");
                        }
                    }
                    else if (args.Length == 2)
                    {
                        // L = 0      L = 1             L=2    
                        // CmdletJack <C:\Path\Cmdlets> search (or -a)
                        //            arg1              arg2
                        // user included the -a option, or the directory with 'search' but no search pattern

                        if (arg2.ToLower() == "search")
                        {
                            Console.WriteLine("Must specify a seach pattern.");
                        }
                        else
                        {
                            if (arg2 == "-a" || arg2 == "/a" || arg2 == "-A" || arg2 == "/A")
                            {
                                append = true;
                                if (Directory.Exists(arg1) || File.Exists(arg1))
                                {
                                    StartJob(arg1);
                                }
                                else
                                {
                                    Console.WriteLine("Invalid directory or file specification.");
                                }
                            }
                            else
                            {
                                Console.WriteLine("Invalid arguments.");
                            }
                        }
                    }
                    else if (args.Length == 3)
                    {
                        // L = 0      L = 1             L=2    L=3       
                        // CmdletJack <C:\Path\Cmdlets> search <pattern> 
                        //            arg1              arg2   arg3
                        //User provided a path, specified 'search', and search pattern.

                        if (arg2.ToLower() != "search")
                        {
                            Console.WriteLine("Invalid arguments.");
                        }
                        else
                        {
                            searchPattern = arg3;
                            if (Directory.Exists(arg1) || File.Exists(arg1))
                            {
                                DoSearch(arg1);
                            }
                            else
                            {
                                Console.WriteLine("Invalid directory or file specification.");
                            }
                        }
                    }

                    else if (args.Length == 4)
                    {
                        // L = 0      L = 1             L=2    L=3       L=4 
                        // CmdletJack <C:\Path\Cmdlets> search <pattern> [<matchcase|wholeword|both>
                        //            arg1              arg2   arg3      arg4

                        if (Directory.Exists(arg1) || File.Exists(arg1))
                        {
                            if (arg2.ToLower() != "search")
                            {
                                Console.WriteLine("Invalid arguments.");
                            }
                            else if (arg4 == "-a" || arg4 == "/a" || arg4 == "-A" || arg4 == "/A")
                            {
                                Console.WriteLine("Sorry, the append option is not available for searches.");
                            }
                            else
                            {
                                searchPattern = arg3;
                                if (arg4.ToLower() == "matchcase")
                                {
                                    matchcase = true;
                                    wholeword = false;
                                }
                                else if (arg4.ToLower() == "wholeword")
                                {
                                    matchcase = false;
                                    wholeword = true;
                                }
                                else if (arg4.ToLower() == "both")
                                {
                                    matchcase = true;
                                    wholeword = true;
                                }
                                if (Directory.Exists(arg1) || File.Exists(arg1))
                                {
                                    DoSearch(arg1);
                                }
                            }
                        }
                        else
                        {
                            Console.WriteLine("Invalid directory or file specification.");
                            throw new DirectoryNotFoundException();
                        }
                    }
                }
            }

            catch (IndexOutOfRangeException IXex)
            {
                Console.WriteLine(IXex.Message);
            }
            catch (ArgumentException ArgEx)
            {
                Console.WriteLine(ArgEx.Message);
            }
            catch (UnauthorizedAccessException UAEx)
            {
                Console.WriteLine("Main method: UnauthorizedAccessException - {0}", UAEx.StackTrace);
            }
            catch (PathTooLongException PTLEx)
            {
                Console.WriteLine(PTLEx.Message);
            }
            catch (FileNotFoundException nfe)
            {
                Console.WriteLine(nfe.Message);
            }
            catch (NotSupportedException NSEx)
            {
                Console.WriteLine(NSEx.Message);
            }
            catch (NullReferenceException NullEx)
            {
                Console.WriteLine(NullEx.Message);
            }
            catch (DirectoryNotFoundException DirNFEx)
            {
                Console.WriteLine(DirNFEx.Message);
            }
            catch (XmlException)
            {
                Console.WriteLine("One of the XML files in the specified path is corrupt.\nTry downloading the file(s) from the server again.");
            }
        }

        private static void StartJob(string path)
        {
            // Called from Main when a path is provided to run reports

            fi = new FileInfo(path);
            string ext = String.Empty;
            ext = Path.GetExtension(fi.FullName);
            string lastjob = String.Empty;
            if (ext == String.Empty)
            {
                if (Directory.Exists(fi.FullName))
                {
                    saveDir = fi.FullName;
                    DirectoryInfo di = new DirectoryInfo(fi.FullName);
                    IEnumerable<string> projfiles = Directory.EnumerateFiles(di.FullName, "*.xml", SearchOption.AllDirectories);
                    JobList = new List<string>(projfiles);
                    ValidateJobList();
                    int c = JobList.Count;
                    lastjob = JobList[c - 1];
                }
            }
            else
            {
                JobList = new List<string>();
                saveDir = fi.DirectoryName;
                projName = Path.GetFileNameWithoutExtension(fi.FullName);
                JobList.Add(fi.FullName);
                ValidateJobList();
                lastjob = fi.FullName;
            }

            foreach (string proj in JobList)
            {
                Tallies = new CmdletMetrics();
                projName = Path.GetFileNameWithoutExtension(proj);
                Tallies.CmdletProject = projName;

                if (lastjob == proj)
                {
                    allProjsFinished = true;
                }

                DoReports(proj);

            }
            WriteWarnings();
        }

        private static void ValidateJobList()
        {
            // Confirms that the psmaml XML files can be processed.  
            // Any other XML files are ommitted from the list.
            List<string> goodList = new List<string>();
            string job = String.Empty;
            try
            {
                foreach (string j in JobList)
                {
                    FileInfo fi = new FileInfo(j);
                    if (File.Exists(fi.FullName))
                    {
                        job = fi.Name;
                        using (StreamReader sr = new StreamReader(fi.FullName))
                        {
                            string xmlLine = sr.ReadLine();
                            if (xmlLine.Substring(0, 5) == "<?xml")
                            {
                                xd = XDocument.Load(File.OpenRead(fi.FullName));
                                IEnumerable<XElement> commands = xd.Descendants(nsCMD + "command");
                                int cCount = commands.Count<XElement>();
                                if (cCount == 0)
                                {
                                    Console.WriteLine("{0} is an empty project and will not be processed.", fi.Name);
                                }
                                else if (cCount > 0)
                                {
                                    goodList.Add(j);
                                }
                            }
                            else
                            {
                                Console.WriteLine("{0} is an invalid cmdlet xml file and will not be processed.", fi.Name);
                            }
                        }
                    }
                }
                JobList.Clear();
                JobList = goodList;
            }
            catch (InvalidOperationException invalEx)
            {
                string msg = invalEx.StackTrace;
                CreateWarning(job, "N/A", "N/A", "Method", "ValidateJobList", "InvalidOperationException - " + msg, "Internal");
            }
            catch (XmlException xmlEx)
            {
                string msg = xmlEx.Message;
                Console.WriteLine("{0} is an invalid or corrupt xml file.\n{1}", job, msg);
            }


        }

        private static void DoReports(string projfile)
        {
            try
            {
                // Called from the StartJob method.
                // Process XML file and initialize the CmdletData class.
                // Uses the psConvCSV.Invoke method for output to a spreadsheet.

                xd = XDocument.Load(File.OpenRead(projfile));
                IEnumerable<XElement> commands = xd.Root.Descendants(nsCMD + "command");
                Tallies.CmdletCount = commands.Count<XElement>();
                ProjectCmdlets.Clear();

                foreach (XElement cmd in commands)
                {
                    CmdletData ProjCmds = new CmdletData();
                    GatherData(ProjCmds, cmd);
                    ProjCmds.Project = projName;
                    ProjectCmdlets.Add(ProjCmds);
                }

                PowerShell psConvCSV = PowerShell.Create();
                psConvCSV.AddCommand("ConvertTo-Csv");
                psConvCSV.AddParameter("NoTypeInformation");


                string cvspath = String.Empty;

                PrepDescrData();
                string csvPath = Path.Combine(saveDir, "CJ_Descriptions.csv");
                if (allProjsFinished)
                {
                    if (DescrList.Count > 0)
                    {
                        IEnumerable<string> csvStrings = (IEnumerable<string>)psConvCSV.Invoke<string>(DescrList);
                        WriteReport(csvPath, csvStrings, "CmdletDescrData");
                    }
                }

                PrepParamData();
                csvPath = Path.Combine(saveDir, "CJ_Parameters.csv");
                if (allProjsFinished)
                {
                    if (ParamsList.Count > 0)
                    {
                        IEnumerable<string> csvStrings = (IEnumerable<string>)psConvCSV.Invoke<string>(ParamsList);
                        WriteReport(csvPath, csvStrings, "CmdletParamData");
                    }
                }

                PrepExampleData();
                csvPath = Path.Combine(saveDir, "CJ_Examples.csv");
                if (allProjsFinished)
                {
                    if (ExampleList.Count > 0)
                    {
                        IEnumerable<string> csvStrings = (IEnumerable<string>)psConvCSV.Invoke<string>(ExampleList);
                        WriteReport(csvPath, csvStrings, "CmdletExampleData");
                    }
                }

                PrepLinkData();
                csvPath = Path.Combine(saveDir, "CJ_Links.csv");
                if (allProjsFinished)
                {
                    IEnumerable<string> csvStrings = (IEnumerable<string>)psConvCSV.Invoke<string>(LinksList);
                    WriteReport(csvPath, csvStrings, "CmdletLinkData");
                }


                PrepInputOutputData();
                csvPath = Path.Combine(saveDir, "CJ_InOut.csv");
                if (allProjsFinished)
                {
                    IEnumerable<string> csvStrings = (IEnumerable<string>)psConvCSV.Invoke<string>(InOutList);
                    WriteReport(csvPath, csvStrings, "CmdletInputOutputData");
                }


                PrepSummaryData();
                csvPath = Path.Combine(saveDir, "CJ_Summary.csv");
                if (allProjsFinished)
                {
                    if (TalliesList.Count > 0)
                    {
                        IEnumerable<string> csvStrings = (IEnumerable<string>)psConvCSV.Invoke<string>(SummaryList);
                        WriteReport(csvPath, csvStrings, "Summary");
                    }
                }

            }
            catch (FileNotFoundException FNFEx)
            {
                string msg = FNFEx.Message + " " + FNFEx.StackTrace;
                CreateWarning(projName, "N/A", "N/A", "Method", "DoProcess", "FileNotFoundException - " + msg, "Internal");
                Console.WriteLine("{0} not found.", projfile);
            }
            catch (NullReferenceException NullEx)
            {
                string msg = NullEx.StackTrace;
                CreateWarning(projName, "N/A", "N/A", "Method", "DoProcess", "NullReferenceException - " + msg, "Internal");
            }
            catch (IOException IOEx)
            {
                string msg = IOEx.Message + " " + IOEx.StackTrace;
                CreateWarning(projName, "N/A", "N/A", "Method", "DoProcess", "IOException - " + msg, "Internal");
            }

        }

        private static void DoSearch(string projfile)
        {
            // Called from Main
            try
            {

                if (Directory.Exists(projfile))
                {
                    IEnumerable<string> projfiles = Directory.EnumerateFiles(projfile, "*.xml");
                    foreach (string p in projfiles)
                    {
                        SearchProj(p);
                    }
                }
                else if (File.Exists(projfile))
                {
                    SearchProj(projfile);
                }

                switch (curSearchHits)
                {
                    case 0:
                        Console.WriteLine("Found no occurences of: {0}", searchPattern);
                        break;
                    case 1:
                        Console.WriteLine("Found one occurence of: {0}", searchPattern);
                        break;
                    default:
                        Console.WriteLine("Found {0} occurences of: {1}", curSearchHits.ToString(), searchPattern);
                        break;
                }

            }
            catch (UnauthorizedAccessException UAEx)
            {
                Console.WriteLine("{0}\nDo you have any of the XML files open?",UAEx.Message);
            }
        }
        
        private static void SearchProj(string projfile)
        {
            // Called from the DoSearch method

            Tallies = new CmdletMetrics(); 
            xd = XDocument.Load(File.OpenRead(projfile));
            IEnumerable<XElement> commands = xd.Root.Descendants(nsCMD + "command");
            int cc = commands.Count<XElement>();
            Tallies.CmdletCount = cc;
            ProjectCmdlets.Clear();

            FileInfo fi = new FileInfo(projfile);
            saveDir = fi.DirectoryName;
            projName = Path.GetFileNameWithoutExtension(fi.FullName);
            try
            {
                foreach (XElement cmd in commands)
                {
                    CmdletData ProjCmds = new CmdletData();
                    ProjCmds.Project = projName;
                    GatherData(ProjCmds, cmd);
                    ProjectCmdlets.Add(ProjCmds);
                }

                PowerShell psConvCSV = PowerShell.Create();
                psConvCSV.AddCommand("ConvertTo-Csv");
                psConvCSV.AddParameter("NoTypeInformation");

                string srchOpt = String.Empty;
                if (matchcase == true && wholeword == true)
                {
                    PrepSearchData(true, true);
                    srchOpt = "MatchCase & WholeWord";
                }
                else if (matchcase == true && wholeword == false)
                {
                    PrepSearchData(true, false);
                    srchOpt = "MatchCase";
                }
                else if (matchcase == false && wholeword == true)
                {
                    PrepSearchData(false, true);
                    srchOpt = "WholeWord";
                }
                if (matchcase == false && wholeword == false)
                {
                    PrepSearchData(false, false);
                    srchOpt = "None";
                }

                string csvPath = Path.Combine(saveDir, "CJ_SearchResults.csv");
                if (Tallies.SearchHit > 0)
                {
                    IEnumerable<string> csvStrings = (IEnumerable<string>)psConvCSV.Invoke<string>(SearchResultList);

                    curSearchHits = Tallies.SearchHit;
                    WriteReport(csvPath, csvStrings, "SearchResults");
                }

            }
            catch (UnauthorizedAccessException UAEx)
            {
                 Console.WriteLine("{0}\nDo you have any of the XML files open? (SearchProj)",UAEx.Message);
            }
            catch (ArgumentNullException NullEx)
            {
                Console.WriteLine(NullEx.Message);
            }             
        }

        private static bool Find(string searchTxt, string option)
        {
            // Called from PrepSearchData do perform the Regex expressions
            bool found = false;

            if (option == "Both")
            {
                MatchCollection hits = Regex.Matches(searchTxt, "\\b" + searchPattern + "\\b");
                if (hits.Count > 0)
                {
                    found = true;
                }
            }
            else if (option == "MatchCase")
            {
                MatchCollection hits = Regex.Matches(searchTxt, searchPattern);
                if (hits.Count > 0)
                {
                    found = true;
                }
            }
            else if (option == "WholeWord")
            {
                MatchCollection hits = Regex.Matches(searchTxt, "\\b" + searchPattern + "\\b", RegexOptions.IgnoreCase);
                if (hits.Count > 0)
                {
                    found = true;
                }
            }
            else if (option == "None")
            {
                MatchCollection hits = Regex.Matches(searchTxt, searchPattern, RegexOptions.IgnoreCase);
                if (hits.Count > 0)
                {
                    found = true;
                }
            }
            return found;

        }

        private static CmdletData GatherData(CmdletData ProjCmds, XElement cmd)
        {
            // This is the method that parses the XML files, 
            // called from the DoReports and SearchProj methods.
            // Populates objects with element values from the XML file.

            string name = String.Empty;
            string verb = String.Empty;
            string noun = String.Empty;
            string category = String.Empty;

            try
            {
                // Get Synopsis and Description
                name = cmd.Element(nsCMD + "details").Element(nsCMD + "name").Value.ToString();
                string synopsis = cmd.Element(nsCMD + "details").Element(nsMML + "description").Value.ToString();
                string description = cmd.Element(nsMML + "description").Value.ToString();
                category = "Gather Data - Synopsis";

                ProjCmds.Name = name;
                ProjCmds.Synopsis = synopsis;
                ProjCmds.Description = description;

                SplitVerb(name, out verb, out noun);

                // note missing
                if (synopsis == String.Empty)
                {
                    CreateWarning(projName, verb, noun, "Synopsis", "Synopsis - missing", "need to write", "QA");
                    Tallies.NoSynopsis++;
                }
                else
                {
                    QAVerbStart(synopsis, verb, noun, "Synopsis", "");
                }
                category = "Gather Data - Descr";
                if (description == String.Empty)
                {
                    CreateWarning(projName, verb, noun, "Description", "Description - missing", "need to write", "QA");
                    Tallies.NoDescription++;
                }
                else
                {
                    // note starts correctly
                    QADescrStart(description, name, verb, noun);
                }

                // Get Parameters
                IEnumerable<XElement> CmdParams = cmd.Element(nsCMD + "parameters").Elements(nsCMD + "parameter");

                // temp dictionaries to use to set CmdletData (ProjCmds object)
                Dictionary<string, string> CmdletParams = new Dictionary<string, string>();
                Dictionary<string, string> CmdletParamsTypes = new Dictionary<string, string>();

                category = "GatherData - Params";
                string pName = String.Empty;
                string pDescr = String.Empty;
                string pType = String.Empty;

                foreach (XElement xp in CmdParams)
                {
                    if (xp.Element(nsMML + "name").IsEmpty || string.IsNullOrWhiteSpace(xp.Element(nsMML + "name").Value))
                    {
                        CreateWarning(projName, verb, noun, "Parameter", "No content", "Empty Element", "Element");
                    }
                    else
                    {
                        pName = xp.Element(nsMML + "name").Value.ToString();
                        pDescr = xp.Element(nsMML + "description").Value.ToString();
                        pType = xp.Element(nsDEV + "type").Element(nsMML + "name").Value.ToString();
                        if (pType == "" || pType == null)
                        {
                            pType = "Unknown";
                            CreateWarning(projName, verb, noun, "Parameter", "Unknown data type", pName, "Element");
                        }
                        bool missingPDescr = false;

                        //note missing
                        if (pDescr == String.Empty)
                        {
                            Tallies.NoParameter++;
                            missingPDescr = true;
                            CreateWarning(projName, verb, noun, "Parameter", "Param descr - missing", pName, "QA");
                        }
                        else
                        {
                            if (!missingPDescr)
                            {
                                // note verb QA
                                QAVerbStart(pDescr, verb, noun, "Parameter", pName);
                            }
                        }

                        try
                        {
                            CmdletParams.Add(pName, pDescr);

                            // For unique keys for the CmdetParamTypes dictionary,
                            // prepend with the cmdlet name and a dot.

                            string pNameT = name + "." + pName;
                            CmdletParamsTypes.Add(pNameT, pType);
                        }

                        catch (NullReferenceException nullEx)
                        {
                            string msg = nullEx.StackTrace;
                            CreateWarning(projName, verb, noun, "Parameter", pType, "NullRefrenceException - " + msg, "Internal");


                        }
                        catch (ArgumentNullException argNullEx)
                        {
                            string msg = argNullEx.Message + " " + argNullEx.StackTrace;
                            CreateWarning(projName, verb, noun, "Paramter", pType, "ArgumentNullExcpeiton - " + msg, "Internal");
                        }

                        catch (ArgumentException argEx)
                        {
                            string msg = argEx.Message + " " + argEx.StackTrace;
                            CreateWarning(projName, verb, noun, "Parameter", pType, "ArgumentException - " + msg, "Internal");
                        }
                    }
                }
                ProjCmds.Params = CmdletParams;
                ProjCmds.ParamTypes = CmdletParamsTypes;

                // Get inputTypes and outputTypes (Return Values)

                IEnumerable<XElement> CmdInputTypes = cmd.Element(nsCMD + "inputTypes").Elements(nsCMD + "inputType");
                IEnumerable<XElement> CmdOutputTypes = cmd.Element(nsCMD + "returnValues").Elements(nsCMD + "returnValue");

                // Temp dictionaries to set ProjCmds properties
                Dictionary<string, string> inputTD = new Dictionary<string, string>();
                Dictionary<string, string> inputTU = new Dictionary<string, string>();
                Dictionary<string, string> outputTD = new Dictionary<string, string>();
                Dictionary<string, string> outputTU = new Dictionary<string, string>();


                foreach (XElement xp in CmdInputTypes)
                {
                    string inT = String.Empty;
                    string inD = String.Empty;
                    string inU = String.Empty;


                    if (xp.Element(nsDEV + "type").Element(nsMML + "name").IsEmpty || string.IsNullOrWhiteSpace(xp.Element(nsDEV + "type").Element(nsMML + "name").Value))
                    {
                        CreateWarning(projName, verb, noun, "Input Type Name", "No content", "Empty Element", "Element");
                        Tallies.NoInputObj++;
                    }
                    else
                    {
                        inT = xp.Element(nsDEV + "type").Element(nsMML + "name").Value.ToString();
                    }
                    if (xp.Element(nsDEV + "type").Element(nsMML + "uri").IsEmpty || string.IsNullOrWhiteSpace(xp.Element(nsDEV + "type").Element(nsMML + "uri").Value))
                    {
                        CreateWarning(projName, verb, noun, "Input Type URI", "No content", "Empty Element", "Element");
                        Tallies.NoInputObjURI++;
                    }
                    else
                    {
                        inU = xp.Element(nsDEV + "type").Element(nsMML + "uri").Value.ToString();
                    }
                    if (xp.Element(nsMML + "description").IsEmpty || string.IsNullOrWhiteSpace(xp.Element(nsMML + "description").Value))
                    {
                        CreateWarning(projName, verb, noun, "Input Type Description", "No content", "Empty Element", "Element");
                        Tallies.NoInputObjDescr++;
                    }
                    else
                    {
                        inD = xp.Element(nsMML + "description").Value.ToString();

                    }
                    inputTD.Add(inT, inD);
                    inputTU.Add(inT, inU);

                }

                foreach (XElement xp in CmdOutputTypes)
                {

                    string outT = String.Empty;
                    string outD = String.Empty;
                    string outU = String.Empty;

                    if (xp.Element(nsDEV + "type").Element(nsMML + "name").IsEmpty || string.IsNullOrWhiteSpace(xp.Element(nsDEV + "type").Element(nsMML + "name").Value))
                    {
                        CreateWarning(projName, verb, noun, "Output Type Name", "No content", "Empty Element", "Element");
                        Tallies.NoOutputObj++;
                    }
                    else
                    {
                        outT = xp.Element(nsDEV + "type").Element(nsMML + "name").Value.ToString();
                    }
                    if (xp.Element(nsDEV + "type").Element(nsMML + "uri").IsEmpty || string.IsNullOrWhiteSpace(xp.Element(nsDEV + "type").Element(nsMML + "uri").Value))
                    {
                        CreateWarning(projName, verb, noun, "Output Type URI", "No content", "Empty Element", "Element");
                        Tallies.NoOutputObjURI++;
                    }
                    else
                    {
                        outU = xp.Element(nsDEV + "type").Element(nsMML + "uri").Value.ToString();

                    }
                    if (xp.Element(nsMML + "description").IsEmpty || string.IsNullOrWhiteSpace(xp.Element(nsMML + "description").Value))
                    {
                        CreateWarning(projName, verb, noun, "Output Type Description", "No content", "Empty Element", "Element");
                        Tallies.NoOutputObjDescr++;
                    }
                    else
                    {
                        outD = xp.Element(nsMML + "description").Value.ToString();

                    }
                    outputTD.Add(outT, outD);
                    outputTU.Add(outT, outU);
                }
                ProjCmds.InputTDescr = inputTD;
                ProjCmds.InputTUri = inputTU;
                ProjCmds.OutputTDescr = outputTD;
                ProjCmds.OutputTUri = outputTU;


                // Get Examples

                bool hasExCode = false;

                IEnumerable<XElement> CmdExamples = cmd.Element(nsCMD + "examples").Elements(nsCMD + "example");
                Dictionary<string, string> CmdletExampleCodes = new Dictionary<string, string>();
                Dictionary<string, string> CmdletExampleDescriptions = new Dictionary<string, string>();
                int exnum = 1;
                category = "GatherData - Examples";
                foreach (XElement xe in CmdExamples)
                {

                    string eTitle = String.Empty;
                    string eCode = String.Empty;
                    string eDescr = String.Empty;

                    if (xe.Element(nsMML + "title").IsEmpty || string.IsNullOrWhiteSpace(xe.Element(nsMML + "title").Value))
                    {
                        // missing title
                        string guid = Guid.NewGuid().ToString();
                        string g = guid.Substring(0, guid.IndexOf("-"));
                        eTitle = "Missing example title - " + g;
                        CreateWarning(projName, verb, noun, "Example", "missing title", "Example " + exnum.ToString(), "Element");
                    }
                    else
                    {
                        eTitle = xe.Element(nsMML + "title").Value.ToString();
                    }


                    IEnumerable<XElement> elems = xe.Elements();
                    List<string> elemList = new List<string>();
                    foreach (XElement el in elems)
                    {
                        elemList.Add(el.Name.ToString());
                    }

                    if (elemList.Contains("{http://schemas.microsoft.com/maml/dev/2004/10}code"))
                    {

                        if (xe.Element(nsDEV + "code").IsEmpty || string.IsNullOrWhiteSpace(xe.Element(nsDEV + "code").Value))
                        {
                            CreateWarning(projName, verb, noun, "Example", "empty code block", "Example " + exnum.ToString(), "Element");
                        }
                        else
                        {
                            eCode = xe.Element(nsDEV + "code").Value.ToString();
                            // code only has a prompt 'PS C:\>'
                            if (eCode.Length < 11)
                            {
                                CreateWarning(projName, verb, noun, "Example", "Code - only has a prompt", "Example " + exnum.ToString() + ": " + eCode, "QA");
                                Tallies.CodeIssue++;

                            }
                            else
                            {
                                QASniffCode(projName, eCode, verb, noun, exnum);
                                hasExCode = true;
                            }
                        }

                    }
                    else
                    {
                        CreateWarning(projName, verb, noun, "Example", "new schema", "only commandlines", "Schema");

                        foreach (XElement xecl in xe.Element(nsCMD + "commandLines").Elements(nsCMD + "commandLine"))
                        {
                            foreach (XElement xeclt in xecl.Elements("commandText"))
                            {
                                if (xeclt.IsEmpty || string.IsNullOrWhiteSpace(xeclt.Value))
                                {
                                    CreateWarning(projName, verb, noun, "Example", "new schema", "only commandlines", "Schema");
                                }
                                else
                                {
                                    CreateWarning(projName, verb, noun, "Example", "code in commandLines schema", eCode, "Schema");
                                    eCode = xeclt.Value.ToString();
                                    QASniffCode(projName, eCode, verb, noun, exnum);
                                    hasExCode = true;
                                }
                            }
                        }
                    }
                    if (xe.Element(nsDEV + "remarks").IsEmpty || string.IsNullOrWhiteSpace(xe.Element(nsDEV + "remarks").Value))
                    {
                        CreateWarning(projName, verb, noun, "Example", "missing description", "Example " + exnum.ToString(), "Element");
                    }
                    else
                    {
                        eDescr = xe.Element(nsDEV + "remarks").Value.ToString();
                        if (eDescr.Substring(0, 1) == " ")
                        {
                            CreateWarning(projName, verb, noun, "Example", "Description - starts with a space", eDescr.Substring(0, 20) + " ...", "QA");
                        }
                    }

                    try
                    {
                        CmdletExampleCodes.Add(eTitle, eCode);
                    }
                    catch (ArgumentException ArgEx)
                    {
                        string msg = ArgEx.Message + " " + ArgEx.StackTrace;
                        CreateWarning(projName, verb, noun, "Method", category, "ArgumentException - " + msg, "Internal");
                    }
                    try
                    {
                        CmdletExampleDescriptions.Add(eTitle, eDescr);
                    }
                    catch (ArgumentException ArgEx)
                    {
                        string msg = ArgEx.Message + " " + ArgEx.StackTrace;
                        CreateWarning(projName, verb, noun, "Method", category, "ArgumentException - " + msg, "Internal");
                    }
                    exnum++;
                }
                ProjCmds.ExampleCodes = CmdletExampleCodes;
                ProjCmds.ExampleDescriptions = CmdletExampleDescriptions;

                if (CmdExamples.Count<XElement>() == 0 || hasExCode == false)
                {
                    CreateWarning(projName, verb, noun, "Example", "Example - missing", "need to write", "QA");
                    Tallies.NoExample++;
                }

                // Get Links
                IEnumerable<XElement> CmdLinks = cmd.Element(nsMML + "relatedLinks").Elements(nsMML + "navigationLink");
                Dictionary<string, string> CmdletLinks = new Dictionary<string, string>();
                string lTxt = String.Empty;
                category = "GatherData - Links";

                foreach (XElement xl in CmdLinks)
                {

                    if (!xl.Element(nsMML + "linkText").IsEmpty || string.IsNullOrWhiteSpace(xl.Element(nsMML + "linkText").Value))
                    {
                        lTxt = xl.Element(nsMML + "linkText").Value.ToString();

                        if (lTxt == ProjCmds.Name)
                        {
                            CreateWarning(projName, verb, noun, "Link", "Links - topic links to itself", "need to delete", "QA");
                            Tallies.LinkIssue++;
                        }
                    }
                    if (lTxt == String.Empty)
                    {
                        CreateWarning(projName, verb, noun, "Link", "link text", "empty value", "Element");
                        Tallies.LinkIssue++;
                    }



                    string lUri = xl.Element(nsMML + "uri").Value.ToString();
                    try
                    {
                        QASniffURI(projName, lUri, verb, noun);
                        if (lUri.StartsWith("http"))
                        {
                            CmdletLinks.Add(lTxt, lUri);
                        }

                    }

                    catch (ArgumentException)
                    {
                        CreateWarning(projName, verb, noun, "Link", "Links - duplicate link in topic", lTxt, "QA");
                        Tallies.LinkIssue++;
                    }
                }
                ProjCmds.Links = CmdletLinks;

                bool hasFWLink = false;
                foreach (KeyValuePair<string, string> kvp in CmdletLinks)
                {
                    if (kvp.Value.Contains("fwlink"))
                    {
                        hasFWLink = true;
                    }
                }
                if (hasFWLink == false)
                {
                    // add empty row to report if no FWLink
                    CmdletLinks.Add("Online Version:", String.Empty);

                    CreateWarning(projName, verb, noun, "Link", "Links - missing FWLink", "need to create", "QA");
                    Tallies.NoFWLink++;
                }


            }


            catch (NullReferenceException NullEx)
            {
                string msg = NullEx.StackTrace;
                CreateWarning(projName, verb, noun, "Method", category, "NullReferenceException - " + msg, "Internal");
            }

            catch (ArgumentException ArgEx)
            {
                string msg = ArgEx.Message + " " + ArgEx.StackTrace;
                CreateWarning(projName, verb, noun, "Method", category, "ArgumentException - " + msg, "Internal");
            }


            return ProjCmds;

        }

        private static void PrepSearchData(bool match, bool whole)
        {

            string verb = String.Empty;
            string noun = String.Empty;
            string option = String.Empty;

            try
            {
                if (match == true && whole == false)
                {
                    option = "MatchCase";
                }
                else if (match == true && whole == true)
                {
                    option = "Both";
                }
                else if (match == false && whole == true)
                {
                    option = "WholeWord";
                }
                else if (match == false && whole == false)
                {
                    option = "None";
                }

                foreach (CmdletData cd in ProjectCmdlets)
                {
                    projName = cd.Project;
                    SplitVerb(cd.Name, out verb, out noun);

                    //Console.WriteLine("Searching {0}-{1}", verb , noun);

                    if (Find(cd.Synopsis, option))
                    {
                        SearchResults sr = new SearchResults();
                        sr.CmdletProj = projName;
                        sr.SearchText = searchPattern;
                        sr.SearchOption = option;
                        sr.CmdletVerb = verb;
                        sr.CmdletNoun = noun;
                        sr.Element = "Synopsis";
                        sr.Name = String.Empty;
                        sr.Result = cd.Synopsis;
                        sr.TimeStamp = GetTimeStamp();
                        Tallies.SearchHit++;
                        SearchResultList.Add(sr);
                    }
                    if (Find(cd.Description, option))
                    {
                        SearchResults sr = new SearchResults();
                        sr.CmdletProj = projName;
                        sr.SearchText = searchPattern;
                        sr.SearchOption = option;
                        sr.CmdletVerb = verb;
                        sr.CmdletNoun = noun;
                        sr.Element = "Description";
                        sr.Name = String.Empty;
                        sr.Result = cd.Description;
                        sr.TimeStamp = GetTimeStamp();
                        Tallies.SearchHit++;
                        SearchResultList.Add(sr);
                    }
                    if (cd.Params != null)
                    {
                        foreach (KeyValuePair<string, string> kvp in cd.Params)
                        {
                            if (Find(kvp.Value, option))
                            {
                                SearchResults sr = new SearchResults();
                                sr.CmdletProj = projName;
                                sr.SearchText = searchPattern;
                                sr.SearchOption = option;
                                sr.CmdletVerb = verb;
                                sr.CmdletNoun = noun;
                                sr.Element = "Parameter";
                                sr.Name = kvp.Key;
                                sr.Result = kvp.Value;
                                Tallies.SearchHit++;
                                sr.TimeStamp = GetTimeStamp();
                                SearchResultList.Add(sr);
                            }
                        }
                    }

                    if (cd.InputTDescr != null)
                    {
                        foreach (KeyValuePair<string, string> kvp in cd.InputTDescr)
                        {
                            if (Find(kvp.Value, option))
                            {
                                SearchResults sr = new SearchResults();

                                sr.CmdletProj = projName;
                                sr.SearchText = searchPattern;
                                sr.SearchOption = option;
                                sr.CmdletVerb = verb;
                                sr.CmdletNoun = noun;
                                sr.Element = "Input Obj Descr";
                                sr.Name = kvp.Key;
                                sr.Result = kvp.Value;
                                Tallies.SearchHit++;
                                sr.TimeStamp = GetTimeStamp();
                                SearchResultList.Add(sr);
                            }
                        }
                    }

                    if (cd.OutputTDescr != null)
                    {
                        foreach (KeyValuePair<string, string> kvp in cd.OutputTDescr)
                        {
                            if (Find(kvp.Value, option))
                            {
                                SearchResults sr = new SearchResults();

                                sr.CmdletProj = projName;
                                sr.SearchText = searchPattern;
                                sr.SearchOption = option;
                                sr.CmdletVerb = verb;
                                sr.CmdletNoun = noun;
                                sr.Element = "Output Obj Descr";
                                sr.Name = kvp.Key;
                                sr.Result = kvp.Value;
                                Tallies.SearchHit++;
                                sr.TimeStamp = GetTimeStamp();
                                SearchResultList.Add(sr);
                            }
                        }
                    }
                    if (cd.ExampleDescriptions != null)
                    {
                        foreach (KeyValuePair<string, string> kvp in cd.ExampleDescriptions)
                        {
                            if (Find(kvp.Value, option))
                            {
                                SearchResults sr = new SearchResults();
                                sr.CmdletProj = projName;
                                sr.SearchText = searchPattern;
                                sr.SearchOption = option;
                                sr.CmdletVerb = verb;
                                sr.CmdletNoun = noun;
                                sr.Element = "Example Description";
                                sr.Name = kvp.Key;
                                sr.Result = kvp.Value;
                                Tallies.SearchHit++;
                                sr.TimeStamp = GetTimeStamp();
                                SearchResultList.Add(sr);
                            }
                        }
                    }

                    if (cd.ExampleCodes != null)
                    {
                        foreach (KeyValuePair<string, string> kvp in cd.ExampleCodes)
                        {
                            if (Find(kvp.Value, option))
                            {
                                SearchResults sr = new SearchResults();
                                sr.CmdletProj = projName;
                                sr.SearchText = searchPattern;
                                sr.SearchOption = option;
                                sr.CmdletVerb = verb;
                                sr.CmdletNoun = noun;
                                sr.Element = "Example Code";
                                sr.Name = kvp.Key;
                                sr.Result = kvp.Value;
                                Tallies.SearchHit++;
                                sr.TimeStamp = GetTimeStamp();
                                SearchResultList.Add(sr);
                            }
                        }
                    }
                }

            }
            catch (IOException IOEx)
            {
                string msg = IOEx.Message + " " + IOEx.StackTrace;
                CreateWarning(projName, "N/A", "N/A", "Method", "PrepSearchData", "IOException - " + msg, "Internal");
            }
            catch (NullReferenceException NullEx)
            {
                string msg = NullEx.StackTrace;
                CreateWarning(projName, verb, noun, "Method", "PrepSearchData", "NullReferenceException - " + msg, "Internal");
                Console.WriteLine("{0}: {1}-{2}", NullEx.Message, verb, noun);
            }
            catch (ParameterBindingException BEx)
            {
                string msg = BEx.Message + " " + BEx.StackTrace;
                CreateWarning(projName, verb, noun, "Method", "PrepSearchData", "ParameterBindingException - " + msg, "Internal");

            }
            catch (UnauthorizedAccessException UAEx)
            {
                Console.WriteLine(UAEx.Message);
            }
        }

        private static void PrepDescrData()
        {
            try
            {

                foreach (CmdletData cd in ProjectCmdlets)
                {
                    string verb;
                    string noun;
                    SplitVerb(cd.Name, out verb, out noun);

                    CmdletDescrData cdd = new CmdletDescrData();
                    cdd.CmdletProj = projName;
                    cdd.CmdletVerb = verb;
                    cdd.CmdletNoun = noun;
                    cdd.Synopsis = cd.Synopsis;
                    cdd.Description = cd.Description;
                    cdd.TimeStamp = GetTimeStamp();

                    QASniffPowerShellBranding(projName, cd.Synopsis, verb, noun, "Synopsis");
                    QASniffPowerShellBranding(projName, cd.Description, verb, noun, "Description");
                    QASniffCustom(projName, cd.Synopsis, verb, noun, "Synopsis");
                    QASniffCustom(projName, cd.Description, verb, noun, "Description");

                    DescrList.Add(cdd);
                }


            }
            catch (IOException IOEx)
            {
                string msg = IOEx.Message + " " + IOEx.StackTrace;
                CreateWarning(projName, "N/A", "N/A", "Method", "SaveDescrData", "IOException - " + msg, "Internal");

            }
            catch (NullReferenceException NullEx)
            {
                string msg = NullEx.StackTrace;
                CreateWarning(projName, "N/A", "N/A", "Method", "SaveDescrData", "NullReferenceException - " + msg, "Internal");
            }
            catch (ParameterBindingException BEx)
            {
                string msg = BEx.Message + " " + BEx.StackTrace;
                CreateWarning(projName, "N/A", "N/A", "Method", "SaveDescrData", "ParameterBindingException - " + msg, "Internal");


            }


        }

        private static void PrepParamData()
        {
            try
            {

                foreach (CmdletData cd in ProjectCmdlets)
                {
                    string verb;
                    string noun;
                    SplitVerb(cd.Name, out verb, out noun);

                    if (cd.Params.Count == 0)
                    {
                        CreateWarning(projName, verb, noun, "Parameter", "no parameters", "empty element", "Element");

                    }
                    else
                    {
                        foreach (KeyValuePair<string, string> kvp in cd.Params)
                        {
                            CmdletParamData cpd = new CmdletParamData();
                            cpd.CmdletProj = projName;
                            cpd.ParameterName = kvp.Key;
                            cpd.CmdletVerb = verb;
                            cpd.CmdletNoun = noun;
                            cpd.ParameterDescription = kvp.Value;

                            QASniffPowerShellBranding(projName, kvp.Value, verb, noun, "Parameter");
                            QASniffCustom(projName, kvp.Value, verb, noun, "Parameter");

                            string pt = String.Empty;

                            // The paramter name in the ParamTypes dictionary
                            // is prepended with the cmdlet name and a dot.
                            cd.ParamTypes.TryGetValue(cd.Name + "." + kvp.Key, out pt);
                            cpd.ParameterType = pt;
                            if (CommonParams.Contains(kvp.Key))
                            {
                                cpd.Common = "Yes";
                            }
                            else
                            {
                                cpd.Common = "No";
                            }
                            cpd.TimeStamp = GetTimeStamp();
                            ParamsList.Add(cpd);

                        }
                    }
                }

            }
            catch (IOException IOEx)
            {
                string msg = IOEx.Message + " " + IOEx.StackTrace;
                CreateWarning(projName, "N/A", "N/A", "Method", "SaveParamData", "IOException - " + msg, "Internal");
            }
            catch (NullReferenceException NullEx)
            {
                string msg = NullEx.StackTrace;
                CreateWarning(projName, "N/A", "N/A", "Method", "SaveParamData", "NullReferenceException - " + msg, "Internal");
            }
            catch (ParameterBindingException BEx)
            {
                string msg = BEx.Message + " " + BEx.StackTrace;
                CreateWarning(projName, "N/A", "N/A", "Method", "SaveParamData", "ParameterBindingException - " + msg, "Internal");

            }
        }

        private static void PrepExampleData()
        {
            NameHitList.Clear();
            try
            {
                foreach (CmdletData cd in ProjectCmdlets)
                {
                    if (cd.ExampleCodes != null)
                    {
                        string verb;
                        string noun;
                        SplitVerb(cd.Name, out verb, out noun);

                        if (cd.ExampleCodes.Count == 0)
                        {
                            CreateWarning(projName, verb, noun, "Example", "no examples", "empty element", "Element");
                        }
                        else
                        {
                            foreach (KeyValuePair<string, string> kvp in cd.ExampleCodes)
                            {
                                CmdletExampleData ced = new CmdletExampleData();
                                ced.CmdletProj = projName;
                                ced.CmdletVerb = verb;
                                ced.CmdletNoun = noun;
                                ced.ExampleTitle = kvp.Key;
                                ced.ExampleCode = kvp.Value;
                                ced.TimeStamp = GetTimeStamp();

                                string ed = String.Empty;
                                if (cd.ExampleDescriptions.TryGetValue(kvp.Key, out ed))
                                {
                                    ced.ExampleDescr = ed;
                                    QASniffPowerShellBranding(projName, ed, verb, noun, "Example Description");
                                    QASniffCustom(projName, ed, verb, noun, "Example Description");

                                }
                                else
                                {
                                    CreateWarning(projName, verb, noun, "Example", "Failed to get code description", kvp.Key, "Element");
                                }
                                try
                                {
                                    ExampleList.Add(ced);
                                }
                                catch (ArgumentException)
                                {
                                    CreateWarning(projName, verb, noun, "Example", "Failed to get example", kvp.Key, "Element");

                                }
                            }
                        }
                    }
                    else
                    {
                        CreateWarning(projName, "N/A", "N/A", "Example", cd.Name, "No examples", "Element");
                    }
                }
            }
            catch (NullReferenceException NullEx)
            {
                string msg = NullEx.StackTrace;
                CreateWarning(projName, "N/A", "N/A", "Method", "PrepExampleData", "NullReferenceException - " + msg, "Internal");
            }

        }

        private static void PrepInputOutputData()
        {
            try
            {

                foreach (CmdletData cd in ProjectCmdlets)
                {
                    string verb;
                    string noun;
                    SplitVerb(cd.Name, out verb, out noun);

                    foreach (KeyValuePair<string, string> kvp in cd.InputTDescr)
                    {
                        CmdletInputOutputData ciod = new CmdletInputOutputData();
                        ciod.CmdletProj = projName;
                        ciod.CmdletVerb = verb;
                        ciod.CmdletNoun = noun;
                        ciod.Direction = "Input";
                        ciod.TypeName = kvp.Key;
                        ciod.TypeDescr = kvp.Value;
                        string iUri = String.Empty;
                        bool URIhit = cd.InputTUri.TryGetValue(kvp.Key, out iUri);
                        if (URIhit == false)
                        {
                            iUri = String.Empty;
                        }
                        ciod.TypeURI = iUri;
                        ciod.TimeStamp = GetTimeStamp();
                        InOutList.Add(ciod);
                    }
                    foreach (KeyValuePair<string, string> kvp in cd.OutputTDescr)
                    {

                        CmdletInputOutputData ciod = new CmdletInputOutputData();
                        ciod.CmdletProj = projName;
                        ciod.CmdletVerb = verb;
                        ciod.CmdletNoun = noun;
                        ciod.Direction = "Output";
                        ciod.TypeName = kvp.Key;
                        ciod.TypeDescr = kvp.Value;
                        string oUri = String.Empty;
                        bool URIhit = cd.OutputTUri.TryGetValue(kvp.Key, out oUri);
                        if (URIhit == false)
                        {
                            oUri = String.Empty;
                        }
                        ciod.TypeURI = oUri;
                        ciod.TimeStamp = GetTimeStamp();
                        InOutList.Add(ciod);
                    }

                }

            }
            catch (NullReferenceException NullEx)
            {
                string msg = NullEx.StackTrace;
                CreateWarning(projName, "N/A", "N/A", "Method", "SaveLinkData", "NullReferenceException - " + msg, "Internal");

            }
            catch (ParameterBindingException BEx)
            {
                string msg = BEx.Message + " " + BEx.StackTrace;
                CreateWarning(projName, "N/A", "N/A", "Method", "SaveLinkData", "ParameterBindingException - " + msg, "Internal");
            }



        }

        private static void PrepLinkData()
        {
            try
            {

                foreach (CmdletData cd in ProjectCmdlets)
                {
                    string verb;
                    string noun;
                    SplitVerb(cd.Name, out verb, out noun);
                    foreach (KeyValuePair<string, string> kvp in cd.Links)
                    {
                        CmdletLinkData cld = new CmdletLinkData();
                        cld.CmdletProj = projName;
                        cld.CmdletVerb = verb;
                        cld.CmdletNoun = noun;
                        cld.LinkText = kvp.Key;
                        cld.LinkUri = kvp.Value;
                        cld.TimeStamp = GetTimeStamp();
                        LinksList.Add(cld);
                    }
                }
            }
            catch (NullReferenceException NullEx)
            {
                string msg = NullEx.StackTrace;
                CreateWarning(projName, "N/A", "N/A", "Method", "SaveLinkData", "NullReferenceException - " + msg, "Internal");

            }
            catch (ParameterBindingException BEx)
            {
                string msg = BEx.Message + " " + BEx.StackTrace;
                CreateWarning(projName, "N/A", "N/A", "Method", "SaveLinkData", "ParameterBindingException - " + msg, "Internal");
            }

        }

        private static void PrepSummaryData()
        {
            SummaryList.Clear();
            Tallies.CmdletProject = projName;
            Tallies.ParamCount = GetParamCount(projName);
            TalliesList.Add(Tallies);
            foreach (CmdletMetrics metric in TalliesList)
            {
                Summary s = new Summary();
                s.Metric = "Cmdlet Count";
                s.Value = metric.CmdletCount.ToString();
                s.CmdletProj = metric.CmdletProject;
                s.TimeStamp = GetTimeStamp();
                SummaryList.Add(s);
            }
            foreach (CmdletMetrics metric in TalliesList)
            {
                Summary s = new Summary();
                s.Metric = "Param Count";
                s.Value = metric.ParamCount.ToString();
                s.CmdletProj = metric.CmdletProject;
                s.TimeStamp = GetTimeStamp(); 
                SummaryList.Add(s);
            }
            foreach (CmdletMetrics metric in TalliesList)
            {
                Summary s = new Summary();
                s.Metric = "No Synopsis";
                s.Value = metric.NoSynopsis.ToString();
                s.CmdletProj = metric.CmdletProject;
                s.TimeStamp = GetTimeStamp();
                SummaryList.Add(s);
            }
            foreach (CmdletMetrics metric in TalliesList)
            {
                Summary s = new Summary();
                s.Metric = "QA Synopsis no verb";
                s.Value = metric.SynopsisQANoVerb.ToString();
                s.CmdletProj = metric.CmdletProject;
                s.TimeStamp = GetTimeStamp();
                SummaryList.Add(s);
            }
            foreach (CmdletMetrics metric in TalliesList)
            {
                Summary s = new Summary();
                s.Metric = "No Description";
                s.Value = metric.NoDescription.ToString();
                s.CmdletProj = metric.CmdletProject;
                s.TimeStamp = GetTimeStamp();
                SummaryList.Add(s);
            }
            foreach (CmdletMetrics metric in TalliesList)
            {
                Summary s = new Summary();
                s.Metric = "QA long Desc start";
                s.Value = metric.CmdletCount.ToString();
                s.CmdletProj = metric.CmdletProject;
                s.TimeStamp = GetTimeStamp();
                SummaryList.Add(s);
            }
            foreach (CmdletMetrics metric in TalliesList)
            {
                Summary s = new Summary();
                s.Metric = "No Parameter descr";
                s.Value = metric.NoParameter.ToString();
                s.CmdletProj = metric.CmdletProject;
                s.TimeStamp = GetTimeStamp();
                SummaryList.Add(s);
            }
            foreach (CmdletMetrics metric in TalliesList)
            {
                Summary s = new Summary();
                s.Metric = "QA Parameter start";
                s.Value = metric.ParameterQANoVerb.ToString();
                s.CmdletProj = metric.CmdletProject;
                s.TimeStamp = GetTimeStamp();
                SummaryList.Add(s);
            }
            foreach (CmdletMetrics metric in TalliesList)
            {
                Summary s = new Summary();
                s.Metric = "No Input Obj";
                s.Value = metric.NoInputObj.ToString();
                s.CmdletProj = metric.CmdletProject;
                s.TimeStamp = GetTimeStamp();
                SummaryList.Add(s);
            }
            foreach (CmdletMetrics metric in TalliesList)
            {
                Summary s = new Summary();
                s.Metric = "No Input Obj descr";
                s.Value = metric.NoInputObjDescr.ToString();
                s.CmdletProj = metric.CmdletProject;
                s.TimeStamp = GetTimeStamp();
                SummaryList.Add(s);
            }
            foreach (CmdletMetrics metric in TalliesList)
            {
                Summary s = new Summary();
                s.Metric = "No Input Obj URI";
                s.Value = metric.NoInputObjURI.ToString();
                s.CmdletProj = metric.CmdletProject;
                s.TimeStamp = GetTimeStamp();
                SummaryList.Add(s);
            }
            foreach (CmdletMetrics metric in TalliesList)
            {
                Summary s = new Summary();
                s.Metric = "No Output Obj";
                s.Value = metric.NoOutputObj.ToString();
                s.CmdletProj = metric.CmdletProject;
                s.TimeStamp = GetTimeStamp();
                SummaryList.Add(s);
            }
            foreach (CmdletMetrics metric in TalliesList)
            {
                Summary s = new Summary();
                s.Metric = "No Ouput Obj descr";
                s.Value = metric.NoOutputObjDescr.ToString();
                s.CmdletProj = metric.CmdletProject;
                s.TimeStamp = GetTimeStamp();
                SummaryList.Add(s);
            }
            foreach (CmdletMetrics metric in TalliesList)
            {
                Summary s = new Summary();
                s.Metric = "No Ouput Obj URI";
                s.Value = metric.NoOutputObjURI.ToString();
                s.CmdletProj = metric.CmdletProject;
                s.TimeStamp = GetTimeStamp();
                SummaryList.Add(s);
            }
            foreach (CmdletMetrics metric in TalliesList)
            {
                Summary s = new Summary();
                s.Metric = "No Example";
                s.Value = metric.NoExample.ToString();
                s.CmdletProj = metric.CmdletProject;
                s.TimeStamp = GetTimeStamp();
                SummaryList.Add(s);
            }
            foreach (CmdletMetrics metric in TalliesList)
            {
                Summary s = new Summary();
                s.Metric = "QA code issue";
                s.Value = metric.CodeIssue.ToString();
                s.CmdletProj = metric.CmdletProject;
                s.TimeStamp = GetTimeStamp();
                SummaryList.Add(s);
            }
            foreach (CmdletMetrics metric in TalliesList)
            {
                Summary s = new Summary();
                s.Metric = "No FWLink";
                s.Value = metric.NoFWLink.ToString();
                s.CmdletProj = metric.CmdletProject;
                s.TimeStamp = GetTimeStamp();
                SummaryList.Add(s);
            }
            foreach (CmdletMetrics metric in TalliesList)
            {
                Summary s = new Summary();
                s.Metric = "Link issue";
                s.Value = metric.LinkIssue.ToString();
                s.CmdletProj = metric.CmdletProject;
                s.TimeStamp = GetTimeStamp();
                SummaryList.Add(s);
            }
            foreach (CmdletMetrics metric in TalliesList)
            {
                Summary s = new Summary();
                s.Metric = "Term issue";
                s.Value = metric.TermIssue.ToString();
                s.CmdletProj = metric.CmdletProject;
                s.TimeStamp = GetTimeStamp();
                SummaryList.Add(s);
            }
        }
        
        private static void WriteReport(string csvPath, IEnumerable<string> csvStrings, string cmdClass)
        {
            // Called from the DoReports method. The method creates spreadsheets.
            string[] csvS = csvStrings.ToArray<string>();
            string[] DataOut;
            string saveMode = String.Empty;
            StreamWriter sw;

            DataOut = csvS;

            try
            {

                if (File.Exists(csvPath) && append == true)
                {
                    // No need for headers
                    DataOut = csvS.Where((val, idx) => idx > 0).ToArray();
                    saveMode = "Appended";
                }
                else if (File.Exists(csvPath) && append == false)
                {
                    saveMode = "Overwrote";
                }
                else if (!File.Exists(csvPath) && append == true)
                {
                    Console.WriteLine("Cannot append to reports that do not exist.");
                    throw new ArgumentException();
                }
                else if (!File.Exists(csvPath))
                {
                    saveMode = "Created";
                }

                string heads;
                HeaderRow.TryGetValue(cmdClass, out heads);
                csvS[0] = heads;

                using (sw = new StreamWriter(csvPath, append))
                {
                    for (int x = 0; x < DataOut.Count<string>(); x++)
                    {
                        sw.WriteLine(DataOut[x]);
                    }
                }

                Console.WriteLine("{0} {1}", saveMode, csvPath);
            }
            catch (IOException IOEx)
            {
                string msg = IOEx.Message + " " + IOEx.StackTrace;
                CreateWarning(projName, "N/A", "N/A", "Method", "WriteReport", "XmlException - " + msg, "Internal");
                Console.WriteLine(IOEx.Message);
            }
            catch (ArgumentException ArgEx)
            {
                Console.WriteLine(ArgEx.Message);

            }

        }

        private static void CreateWarning(string projName, string cmdVerb, string cmdNoun, string element, string category, string context, string warntype)
        {
            Warning warn = new Warning();
            warn.CmdletProj = projName;
            warn.CmdletNoun = cmdNoun;
            warn.CmdletVerb = cmdVerb;
            warn.Element = element;
            warn.Category = category;
            warn.Context = context;
            warn.WarningType = warntype;
            warn.TimeStamp = GetTimeStamp();
            warn.WarningType = warntype;

            Warns.Add(warn);

        }
        
        private static void WriteWarnings()
        {
            try
            {
                string csvPath = Path.Combine(saveDir, "CJ_Warnings.csv");

                PowerShell psConvCSV = PowerShell.Create();
                psConvCSV.AddCommand("ConvertTo-Csv");
                psConvCSV.AddParameter("NoTypeInformation");

                IEnumerable<string> csvStrings = (IEnumerable<string>)psConvCSV.Invoke<string>(Warns);
                string[] csvS = csvStrings.ToArray<string>();
                string[] DataOut;
                DataOut = csvS;

                string saveMode = String.Empty;

                if (File.Exists(csvPath) && append == true)
                {
                    // No need for headers
                    DataOut = csvS.Where((val, idx) => idx > 0).ToArray();
                    saveMode = "Appended";
                }
                else if (File.Exists(csvPath) && append == false)
                {
                    saveMode = "Overwrote";
                }
                else if (!File.Exists(csvPath) && append == true)
                {
                    Console.WriteLine("Cannot append to reports that do not exist.");
                    throw new ArgumentException();
                }
                else if (!File.Exists(csvPath))
                {
                    saveMode = "Created";
                }

                string heads;
                HeaderRow.TryGetValue("Warning", out heads);
                csvS[0] = heads;


                // Get warning totals
                IEnumerable<Warning> qaQ =
                from warn in Warns
                where warn.WarningType == "QA"
                select warn;

                string qattl = qaQ.Count<Warning>().ToString();

                IEnumerable<Warning> qaCQ =
                from warn in Warns
                where warn.WarningType == "CustomQA"
                select warn;

                string qaqttl = qaCQ.Count<Warning>().ToString();


                IEnumerable<Warning> elementQ =
                from warn in Warns
                where warn.WarningType == "Element"
                select warn;

                string elemttl = elementQ.Count<Warning>().ToString();

                IEnumerable<Warning> internalQ =
                from warn in Warns
                where warn.WarningType == "Internal"
                select warn;

                string intlttl = internalQ.Count<Warning>().ToString();

                IEnumerable<Warning> schemaQ =
                from warn in Warns
                where warn.WarningType == "Schema"
                select warn;

                string schemattl = schemaQ.Count<Warning>().ToString();

                StreamWriter sw;

                if (append == false)
                {
                    using (sw = new StreamWriter(csvPath, false))
                    {
                        for (int x = 0; x < DataOut.Count<string>(); x++)
                        {
                            sw.WriteLine(DataOut[x]);
                        }
                    }
                }
                else
                {
                    using (sw = new StreamWriter(csvPath, true))
                    {
                        for (int x = 0; x < DataOut.Count<string>(); x++)
                        {
                            sw.WriteLine(DataOut[x]);
                        }
                    }
                }


                Console.WriteLine("{0} {1}", saveMode, csvPath);

                Console.WriteLine();
                Console.WriteLine("Logged {0} QA warnings.", qattl);
                Console.WriteLine("Logged {0} Custom QA warnings.", qaqttl);
                Console.WriteLine("Logged {0} Element warnings.", elemttl);
                Console.WriteLine("Logged {0} Schema warnings.", schemattl);
                Console.WriteLine("Logged {0} Internal warnings.", intlttl);
            }
            catch (ArgumentException ArgEx)
            {
                Console.WriteLine("IOException - " + fi.Name + ": " + ArgEx.Message);

            }
            catch (IOException IOEx)
            {
                Console.WriteLine("Cannot save warnings file.");
                Console.WriteLine("IOException - " + fi.Name + ": " + IOEx.Message);

            }
            catch (NullReferenceException NullEx)
            {
                Console.WriteLine("Cannot save warnings file.");
                Console.WriteLine("NullReferenceException - " + fi.Name + ": " + NullEx.Message);

            }
            catch (ParameterBindingException BEx)
            {
                Console.WriteLine("Cannot save warnings file.");
                Console.WriteLine("ParamterBindingException - " + fi.Name + ": " + BEx.Message);

            }


        }

        #endregion

        #region Helper methods

        private static void FillCommonParameters()
        {
            CommonParams.Add("Debug");
            CommonParams.Add("ErrorAction");
            CommonParams.Add("ErrorVariable");
            CommonParams.Add("OutVariable");
            CommonParams.Add("OutBuffer");
            CommonParams.Add("Verbose");
            CommonParams.Add("WarningAction");
            CommonParams.Add("WarningVariable");
            CommonParams.Add("WhatIf");
            CommonParams.Add("Confirm");

        }

        private static void SplitVerb(string name, out string verb, out string noun)
        {
            // Splits the cmdlet verb from the noun.
            int dash = name.IndexOf("-");
            verb = name.Substring(0, dash);
            noun = name.Substring(dash + 1);
        }

        private static int GetParamCount(string proj)
        {

            FileInfo fi = new FileInfo(proj);
            string projName = Path.GetFileNameWithoutExtension(fi.Name);

            IEnumerable<CmdletData> projData =
            from CmdletProj in ProjectCmdlets
            where CmdletProj.Project == projName
            select CmdletProj;

            ProjParamCount = 0;

            foreach (CmdletData cd in projData)
            {
                ProjParamCount = ProjParamCount + cd.Params.Count;
            }
            return ProjParamCount;

        }

        private static bool HasSpecialCaseText(string descr)
        {
            string lowerdescr = descr.ToLower();
            if (lowerdescr.Contains("depricated") ||
                lowerdescr.Contains("do not use") ||
                lowerdescr.Contains("disregard") ||
                lowerdescr.Contains("internal use only"))
            {
                return true;
            }
            else
            {
                return false;
            }

        }

        private static string GetTimeStamp()
        {
            return DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToLongTimeString();
        }

        private static void ShowUsage()
        {
            Console.WriteLine();
            Console.WriteLine("CmdletJack creates spreadsheets (.CSV files) of cmdlet documentation");
            Console.WriteLine("content as reports. Usage is the path to the directory of the cmdlet");
            Console.WriteLine("projects or to a specific file. Reports are automatically generated.");
            Console.WriteLine("Syntax for all projects:");
            Console.WriteLine();
            Console.WriteLine(@"    CmdletJack <C:\Path\Cmdlets>");
            Console.WriteLine();
            Console.WriteLine("Syntax for a specifc project:");
            Console.WriteLine();
            Console.WriteLine(@"    CmdletJack <C:\Path\Cmdlets\VMM.XML>");
            Console.WriteLine();
            Console.WriteLine(@"Existing reports are overwritten, unless you follow with '-a' to append.");
            Console.WriteLine();
            Console.WriteLine("You can also search:");
            Console.WriteLine();
            Console.WriteLine(@"    CmdletJack <C:\Path\Cmdlets> search <pattern> [<matchcase|wholeword|both>]");
            Console.WriteLine();
            Console.WriteLine("Search options: specify 'matchcase' to match case, 'wholeword' to match whole words,");
            Console.WriteLine("or 'both' to use matchcase and wholeword.");
            Console.WriteLine();
            Console.WriteLine("CmdletJack version: 2.5 - (c) 2015 Microsoft Corporation. All rights reserved.");
        }

        private static void DeleteIfExists(string path)
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }

        private static void FillValidReportTypes()
        {
            // but do not include "search"

            ValidReportTypes.Add("descr");
            ValidReportTypes.Add("params");
            ValidReportTypes.Add("paramsnc");
            ValidReportTypes.Add("examples");
            ValidReportTypes.Add("examplesnone");
            ValidReportTypes.Add("links");
            ValidReportTypes.Add("summary");
            ValidReportTypes.Add("all");
        }

        private static void FillApprovedCompanyNames()
        {
            // Source: http://lcaweb/CTP/Trademarks/Pages/FictitiousNames.aspx

            ApprovedCompanyNames.Add("adatum");
            ApprovedCompanyNames.Add("adventure-works");
            ApprovedCompanyNames.Add("alpineskihouse");
            ApprovedCompanyNames.Add("blueyonderairlines");
            ApprovedCompanyNames.Add("cpandl");
            ApprovedCompanyNames.Add("cohovineyard");
            ApprovedCompanyNames.Add("cohowinery");
            ApprovedCompanyNames.Add("cohovineyardandwinery");
            ApprovedCompanyNames.Add("contoso");
            ApprovedCompanyNames.Add("consolidatedmessenger");
            ApprovedCompanyNames.Add("fabrikam");
            ApprovedCompanyNames.Add("fourthcoffee");
            ApprovedCompanyNames.Add("graphicdesigninstitute");
            ApprovedCompanyNames.Add("humongousinsurance");
            ApprovedCompanyNames.Add("litwareinc");
            ApprovedCompanyNames.Add("lucernepublishing");
            ApprovedCompanyNames.Add("margiestravel");
            ApprovedCompanyNames.Add("northwindtraders");
            ApprovedCompanyNames.Add("proseware");
            ApprovedCompanyNames.Add("fineartschool");
            ApprovedCompanyNames.Add("southridgevideo");
            ApprovedCompanyNames.Add("tailspintoys");
            ApprovedCompanyNames.Add("treyresearch");
            ApprovedCompanyNames.Add("thephone-company");
            ApprovedCompanyNames.Add("wideworldimporters");
            ApprovedCompanyNames.Add("wingtiptoys");
            ApprovedCompanyNames.Add("woodgrovebank");
        }

        private static void FillHeaders()
        {
            headersCmdletDescrData.Add("Project");
            headersCmdletDescrData.Add("Verb");
            headersCmdletDescrData.Add("Noun");
            headersCmdletDescrData.Add("Synopsis");
            headersCmdletDescrData.Add("Description");
            headersCmdletDescrData.Add("TimeStamp");
            CreateHeaderRow(headersCmdletDescrData, "CmdletDescrData");

            headersCmdletParamData.Add("Project");
            headersCmdletParamData.Add("Verb");
            headersCmdletParamData.Add("Noun");
            headersCmdletParamData.Add("Param Name");
            headersCmdletParamData.Add("Param Type");
            headersCmdletParamData.Add("Param Descr");
            headersCmdletParamData.Add("Common");
            headersCmdletParamData.Add("TimeStamp");
            CreateHeaderRow(headersCmdletParamData, "CmdletParamData");

            headersCmdletExampleData.Add("Project");
            headersCmdletExampleData.Add("Verb");
            headersCmdletExampleData.Add("Noun");
            headersCmdletExampleData.Add("Title");
            headersCmdletExampleData.Add("Descr");
            headersCmdletExampleData.Add("Code");
            headersCmdletExampleData.Add("TimeStamp");
            CreateHeaderRow(headersCmdletExampleData, "CmdletExampleData");

            headersCmdletInOutData.Add("Project");
            headersCmdletInOutData.Add("Verb");
            headersCmdletInOutData.Add("Noun");
            headersCmdletInOutData.Add("Direction");
            headersCmdletInOutData.Add("Name");
            headersCmdletInOutData.Add("Description");
            headersCmdletInOutData.Add("URI");
            headersCmdletInOutData.Add("TimeStamp");
            CreateHeaderRow(headersCmdletInOutData, "CmdletInputOutputData");

            headersCmdletNameHitsData.Add("Project");
            headersCmdletNameHitsData.Add("Verb");
            headersCmdletNameHitsData.Add("Noun");
            headersCmdletNameHitsData.Add("TimeStamp");
            CreateHeaderRow(headersCmdletNameHitsData, "CmdletNameHitsData");

            headersCmdletLinkData.Add("Project");
            headersCmdletLinkData.Add("Verb");
            headersCmdletLinkData.Add("Noun");
            headersCmdletLinkData.Add("Link Text");
            headersCmdletLinkData.Add("URI");
            headersCmdletLinkData.Add("TimeStamp");
            CreateHeaderRow(headersCmdletLinkData, "CmdletLinkData");

            headersSummary.Add("Metric");
            headersSummary.Add("Value");
            headersSummary.Add("Project");
            headersSummary.Add("TimeStamp");
            CreateHeaderRow(headersSummary, "Summary");

            headersWarning.Add("Project");
            headersWarning.Add("Verb");
            headersWarning.Add("Noun");
            headersWarning.Add("Element");
            headersWarning.Add("Category");
            headersWarning.Add("Context");
            headersWarning.Add("Warning Type");
            headersWarning.Add("TimeStamp");
            CreateHeaderRow(headersWarning, "Warning");

            headersSearchResults.Add("Project");
            headersSearchResults.Add("Search Text");
            headersSearchResults.Add("Search Option");
            headersSearchResults.Add("Verb");
            headersSearchResults.Add("Noun");
            headersSearchResults.Add("Element");
            headersSearchResults.Add("Name");
            headersSearchResults.Add("Occurence");
            headersSearchResults.Add("TimeStamp");
            CreateHeaderRow(headersSearchResults, "SearchResults");

        }

        private static void CreateHeaderRow(List<string> headers, string cmdClass)
        {
            // Creates a sting of the headers in comma separated values
            StringBuilder sb = new StringBuilder();
            foreach (string h in headers)
            {
                sb.Append("\"" + h + "\",");

            }
            char[] charsToTrim = { ',' };
            string heads = sb.ToString().TrimEnd(charsToTrim);
            // Add to the header row dictionary
            HeaderRow.Add(cmdClass, heads);
        }

        #endregion

        #region Quality Assurance detection methods

        private static void QAVerbStart(string descr, string verb, string noun, string element, string pName)
        {
            int sniplen = 20;
            string lowerdescr = descr.ToLower();
            string elem = element.ToLower();


            if (descr.Length < 20)
            {
                sniplen = descr.Length;
            }
            try
            {
                int space = descr.IndexOf(" ");

                if (space == 0)
                {
                    // Description begins with a space
                    if (elem == "synopsis")
                    {
                        CreateWarning(projName, verb, noun, element, "Synopsis - starts with a space", descr.Substring(0, sniplen) + " ...", "QA");
                    }
                    else if (elem == "description")
                    {
                        CreateWarning(projName, verb, noun, element, "Description - starts with a space", pName + ": " + descr.Substring(0, sniplen) + " ...", "QA");
                    }
                }
                else
                {
                    int comma = descr.IndexOf(",");

                    if (descr.Substring(1, 1) == " ")
                    {
                        if (HasSpecialCaseText(descr) == false)
                        {
                            // Sentence starts with a single character word
                            if (pName == String.Empty)
                            {
                                CreateWarning(projName, verb, noun, element, "Synopsis - might not start with a verb", descr.Substring(0, sniplen) + " ...", "QA");
                                Tallies.SynopsisQANoVerb++;
                            }
                            else
                            {
                                CreateWarning(projName, verb, noun, element, "Param descr - might not start with a verb", pName + ": " + descr.Substring(0, sniplen) + " ...", "QA");
                                Tallies.ParameterQANoVerb++;
                            }
                        }
                    }

                    else if (space != -1 && (comma == -1 || space < comma))
                    {
                        if (HasSpecialCaseText(descr) == false)
                        {
                            //Starts with a word longer than one character followed by a space in a sentence with no comma.
                            // -or-
                            //Starts with a word longer than one character followed by a space in a sentence where the comma is further along.
                            if (descr.Substring(space - 1, 1) != "s" && descr.Substring(space - 1, 1) != "y")
                            {
                                // First word does not end with an 's' or 'y'
                                if (pName == String.Empty)
                                {
                                    CreateWarning(projName, verb, noun, element, "Synopsis - might not start with a verb", descr.Substring(0, sniplen) + " ...", "QA");
                                    Tallies.SynopsisQANoVerb++;
                                }
                                else
                                {
                                    CreateWarning(projName, verb, noun, element, "Param descr - might not start with a verb", pName + ": " + descr.Substring(0, sniplen) + " ...", "QA");
                                    Tallies.ParameterQANoVerb++;
                                }
                            }
                        }
                    }

                    else if (comma != -1 && comma < space)
                    {
                        if (HasSpecialCaseText(descr) == false)
                        {
                            //Starts with a word longer than one character followed by a comma
                            if (descr.Substring(comma - 1, 1) != "s" && descr.Substring(comma - 1, 1) != "y")
                            {
                                if (pName == String.Empty)
                                {
                                    CreateWarning(projName, verb, noun, element, "Synopsis - might not start with a verb", descr.Substring(0, sniplen) + " ...", "QA");
                                    Tallies.SynopsisQANoVerb++;
                                }
                                else
                                {
                                    CreateWarning(projName, verb, noun, element, "Param descr - might not start with a verb", pName + ": " + descr.Substring(0, sniplen) + " ...", "QA");
                                    Tallies.ParameterQANoVerb++;
                                }
                            }
                        }
                    }
                }

            }
            catch (ArgumentException ArgEx)
            {
                string msg = ArgEx.Message + " " + ArgEx.StackTrace;
                CreateWarning(projName, verb, noun, "Method", "QAVerbStart", "ArgumentException - " + msg, "Internal");
            }
        }

        private static void QADescrStart(string descr, string cmdlet, string verb, string noun)
        {

            string lowerdescr = descr.ToLower();
            int dlen = descr.Length;
            string good = "the " + cmdlet.ToLower() + " cmdlet ";
            int glen = good.Length;

            try
            {
                if (!lowerdescr.StartsWith(good))
                {
                    if (HasSpecialCaseText(descr) == false)
                    {
                        int l;
                        if (glen > dlen)
                        {
                            l = dlen;
                        }
                        else
                        {
                            l = glen;
                        }

                        CreateWarning(projName, verb, noun, "Description", "Description - might not start correctly", descr.Substring(0, l) + " ...", "QA");
                        Tallies.LongDescrQA++;
                    }
                }
                else
                {
                    if (descr.Length == good.Length)
                    {
                        CreateWarning(projName, verb, noun, "Description", "Description - starts correctly but is incomplete", descr, "QA");
                        Tallies.LongDescrQA++;

                    }
                }
            }
            catch (ArgumentException ArgEx)
            {
                string msg = ArgEx.Message + " " + ArgEx.StackTrace;
                CreateWarning(projName, verb, noun, "Method", "QADescrStart", "ArgumentException - " + msg, "Internal");
            }

        }

        private static void QASniffCode(string projName, string code, string verb, string noun, int exnum)
        {
            string codesnip = String.Empty;


            // no https (SSL layer)
            if (code.Contains("http:"))
            {
                int idx = code.IndexOf("http:");
                codesnip = code.Substring(idx - 5, 30) + " ...";
                CreateWarning(projName, verb, noun, "Example", "Code - http used instead of https", "Example " + exnum.ToString() + ": " + codesnip, "QA");
                Tallies.CodeIssue++;
            }

            // check ficticious company name
            if (code.ToLower().Contains(".com ") || code.Contains(".net ") || code.Contains(".org "))
            {
                int dot = 0;
                dot = code.ToLower().IndexOf(".com ");
                if (dot == -1)
                {
                    dot = code.ToLower().IndexOf(".net ");
                }
                if (dot == -1)
                {
                    dot = code.ToLower().IndexOf(".org ");
                }
                if (dot != -1)
                {
                    string sub1 = code.Substring(0, dot);

                    int slash = sub1.LastIndexOf(@"/");
                    int domdot = sub1.LastIndexOf(".");
                    int mailto = sub1.LastIndexOf("@");
                    int space = sub1.LastIndexOf(" ");
                    int star = sub1.LastIndexOf("*");


                    int[] markers = { slash, domdot, mailto, space, star };
                    Array.Sort(markers);
                    int start = markers[4];
                    codesnip = code.Substring(start + 1, dot - (start + 1));
                    if (codesnip.Substring(0, 1) == "\"")
                    {
                        codesnip = codesnip.Substring(1);
                    }


                    if (!ApprovedCompanyNames.Contains(codesnip.ToLower()))
                    {
                        CreateWarning(projName, verb, noun, "Example", "Code - check company name in URL", "Example " + exnum.ToString() + ": " + codesnip, "QA");
                        Tallies.CodeIssue++;
                    }
                }
            }

        }

        private static void QASniffCustom(string ProjName, string txt, string verb, string noun, string element)
        {
            string qafile = Path.Combine(saveDir, "CustomQA.txt");
            if (File.Exists(qafile))
            {
                foreach (string s in File.ReadLines(qafile))
                {
                    // Make things case-insensitive
                    string term = s.ToLower();
                    string srchTxt = txt.ToLower();
                    if (srchTxt.Contains(term))
                    {
                        CreateWarning(projName, verb, noun, element, "Custom: " + s, txt, "CustomQA");
                    }
                }

            }
        }

        private static void QASniffPowerShellBranding(string ProjName, string txt, string verb, string noun, string element)
        {
            if (txt.Contains(" Powershell "))
            {
                CreateWarning(projName, verb, noun, element, "Branding - wrong PowerShell casing", txt, "QA");
                Tallies.TermIssue++;
            }
            else if (txt.Contains(" PowerShell"))
            {
                int pow = txt.IndexOf(" PowerShell");
                // Windows PowerShell
                string win = txt.Substring(pow - 7, 7);
                if (win != "Windows")
                {
                    CreateWarning(projName, verb, noun, element, "Branding - PowerShell not preceded by Windows", txt, "QA");
                    Tallies.TermIssue++;
                }

            }


        }
        
        private static void QASniffURI(string ProjName, string txt, string verb, string noun)
        {
            //http://go.microsoft.com/fwlink/?LinkID=271361

            bool alphaOK = true;
            bool idOK = true;
            bool idNumeric = true;

            string url = txt.ToLower();

            try
            {
                if (url.StartsWith("http"))
                {
                    if (!url.StartsWith(@"http://go.microsoft.com/fwlink"))
                    {
                        alphaOK = false;
                    }
                    else
                    {
                        int eq = url.LastIndexOf("=");
                        string id = url.Substring(eq + 1);
                        int idLen = id.Length;
                        if (idLen < 5 || id.Contains(" "))
                        {
                            idOK = false;
                        }
                        else
                        {

                            for (int x = 0; x < idLen; x++)
                            {
                                string n = id.Substring(x, 1);
                                if (n == String.Empty || n == null)
                                {
                                    idOK = false;
                                }
                                if (idOK == true)
                                {

                                    int d = Convert.ToInt32(n);
                                    if (d >= 0 && d < 10)
                                    {
                                        idNumeric = true;
                                    }
                                    else
                                    {
                                        idNumeric = false;
                                    }
                                }
                            }
                        }
                    }

                    if (alphaOK == false || idOK == false || idNumeric == false)
                    {
                        CreateWarning(projName, verb, noun, "Link", "Links - malformed FWLink", txt, "QA");
                        Tallies.LinkIssue++;
                    }
                }

            }
            catch (System.FormatException FmtEx)
            {
                string msg = FmtEx.Message + " " + FmtEx.StackTrace;
                CreateWarning(projName, "N/A", "N/A", "Method", "QASniffURI", "FormatException - " + msg, "Internal");
            }



        }

        #endregion

    }

    public class CmdletData
    {
        // This class provides properties for
        // cmdlet documentation elements.

        // The DoReports method initializes this
        // class with all data needed from the xml file.
        public string Project { get; set; }
        public string Name { get; set; }
        public string Synopsis { get; set; }
        public string Description { get; set; }
        public Dictionary<string, string> Params { get; set; }
        public Dictionary<string, string> ParamTypes { get; set; }
        public Dictionary<string, string> ExampleDescriptions { get; set; }
        public Dictionary<string, string> ExampleCodes { get; set; }
        public Dictionary<string, string> Links { get; set; }
        public Dictionary<string, string> InputTDescr { get; set; }
        public Dictionary<string, string> InputTUri {get; set;}
        public Dictionary<string, string> OutputTDescr { get; set; }
        public Dictionary<string, string> OutputTUri {get; set; }

    }

    #region Report classes

    public class CmdletParamData
    {
        public string CmdletProj { get; set; }
        public string CmdletVerb { get; set; }
        public string CmdletNoun { get; set; }
        public string ParameterName { get; set; }
        public string ParameterType { get; set; }
        public string ParameterDescription { get; set; }
        public string Common { get; set; }
        public string TimeStamp { get; set; }
    }
    public class CmdletDescrData
    {
        public string CmdletProj { get; set; }
        public string CmdletVerb { get; set; }
        public string CmdletNoun { get; set; }
        public string Synopsis { get; set; }
        public string Description { get; set; }
        public string TimeStamp { get; set; }
    }
    public class CmdletExampleData
    {
        public string CmdletProj { get; set; }
        public string CmdletVerb { get; set; }
        public string CmdletNoun { get; set; }
        public string ExampleTitle { get; set; }
        public string ExampleDescr { get; set; }
        public string ExampleCode { get; set; }
        public string TimeStamp { get; set; }
    }
    public class CmdletNameHitsData
    {
        public string CmdletProj { get; set; }
        public string CmdletVerb { get; set; }
        public string CmdletNoun { get; set; }
        public string TimeStamp { get; set; }
    }
    public class CmdletInputOutputData
    {
        public string CmdletProj { get; set; }
        public string CmdletVerb { get; set; }
        public string CmdletNoun { get; set; }
        public string Direction {get; set;}
        public string TypeName { get; set; }
        public string TypeDescr { get; set; }
        public string TypeURI { get; set; }
        public string TimeStamp { get; set; }
    }
    public class CmdletNameHitsContextData
    {
        public string CmdletProj { get; set; }
        public string CmdletVerb { get; set; }
        public string CmdletNoun { get; set; }
        public string Context { get; set; }
        public string TimeStamp { get; set; }
    }
    public class CmdletLinkData
    {
        public string CmdletProj { get; set; }
        public string CmdletVerb { get; set; }
        public string CmdletNoun { get; set; }
        public string LinkText { get; set; }
        public string LinkUri { get; set; }
        public string TimeStamp { get; set; }
    }
    public class SearchResults
    {
        public string CmdletProj { get; set; }
        public string SearchText { get; set; }
        public string SearchOption { get; set; }
        public string CmdletVerb { get; set; }
        public string CmdletNoun { get; set; }
        public string Element { get; set; }
        public string Name { get; set; }
        public string Result { get; set; }
        public string TimeStamp { get; set; }
    }
    public class CmdletMetrics
    {
        public string CmdletProject { get; set; }
        public int CmdletCount { get; set; }
        public int ParamCount { get; set; }
        public int NoSynopsis { get; set; }
        public int SynopsisQANoVerb { get; set; }
        public int NoDescription { get; set; }
        public int LongDescrQA { get; set; }
        public int NoParameter { get; set; }
        public int ParameterQANoVerb { get; set; }
        public int NoExample { get; set; }
        public int CodeIssue { get; set; }
        public int NoInputObj { get; set; }
        public int NoInputObjURI { get; set; }
        public int NoInputObjDescr { get; set; }
        public int NoOutputObj { get; set; }
        public int NoOutputObjURI { get; set; }
        public int NoOutputObjDescr { get; set; }
        public int NoFWLink { get; set; }
        public int LinkIssue { get; set; }
        public int TermIssue { get; set; }
        public int SearchHit { get; set; }
        public string TimeStamp { get; set; }
    }
    public class Summary
    {
        public string Metric { get; set; }
        public string Value { get; set; }
        public string CmdletProj { get; set; }
        public string TimeStamp { get; set; }
    }

    public class Warning
    {
        public string CmdletProj { get; set; }
        public string CmdletVerb { get; set; }
        public string CmdletNoun { get; set; }
        public string Element { get; set; }
        public string Category { get; set; }
        public string Context { get; set; }
        public string WarningType { get; set; }
        public string TimeStamp { get; set; }
    }
#endregion
}
