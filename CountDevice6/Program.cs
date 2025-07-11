using System.Text;
//using System.Text.RegularExpressions;
using ClosedXML.Excel;
using ConsoleAppFramework;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using Utf8StringInterpolation;
using ZLogger;
using ZLogger.Providers;

//==
var builder = ConsoleApp.CreateBuilder(args);
builder.ConfigureServices((ctx,services) =>
{
    // Register appconfig.json to IOption<MyConfig>
    services.Configure<MyConfig>(ctx.Configuration);

    // Using Cysharp/ZLogger for logging to file
    services.AddLogging(logging =>
    {
        logging.ClearProviders();
        logging.SetMinimumLevel(LogLevel.Trace);
        var jstTimeZoneInfo = TimeZoneInfo.FindSystemTimeZoneById("Tokyo Standard Time");
        var utcTimeZoneInfo = TimeZoneInfo.Utc;
        logging.AddZLoggerConsole(options =>
        {
            options.UsePlainTextFormatter(formatter => 
            {
                formatter.SetPrefixFormatter($"{0:yyyy-MM-dd'T'HH:mm:sszzz}|{1:short}|", (in MessageTemplate template, in LogInfo info) => template.Format(TimeZoneInfo.ConvertTime(info.Timestamp.Utc, jstTimeZoneInfo), info.LogLevel));
                formatter.SetExceptionFormatter((writer, ex) => Utf8String.Format(writer, $"{ex.Message}"));
            });
        });
        logging.AddZLoggerRollingFile(options =>
        {
            options.UsePlainTextFormatter(formatter => 
            {
                formatter.SetPrefixFormatter($"{0:yyyy-MM-dd'T'HH:mm:sszzz}|{1:short}|", (in MessageTemplate template, in LogInfo info) => template.Format(TimeZoneInfo.ConvertTime(info.Timestamp.Utc, jstTimeZoneInfo), info.LogLevel));
                formatter.SetExceptionFormatter((writer, ex) => Utf8String.Format(writer, $"{ex.Message}"));
            });

            // File name determined by parameters to be rotated
            options.FilePathSelector = (timestamp, sequenceNumber) => $"logs/{timestamp.ToLocalTime():yyyy-MM-dd}_{sequenceNumber:00}.log";
            
            // The period of time for which you want to rotate files at time intervals.
            options.RollingInterval = RollingInterval.Day;
            
            // Limit of size if you want to rotate by file size. (KB)
            options.RollingSizeKB = 1024;        
        });
    });
});

var app = builder.Build();
app.AddCommands<CountDeviceApp>();
app.Run();


public class CountDeviceApp : ConsoleAppBase
{
    bool isAllPass = true;

    readonly ILogger<CountDeviceApp> logger;
    readonly IOptions<MyConfig> config;

    public CountDeviceApp(ILogger<CountDeviceApp> logger,IOptions<MyConfig> config)
    {
        this.logger = logger;
        this.config = config;
    }

    Dictionary<string, MyUsedDevice> MyUsedDevices = new Dictionary<string, MyUsedDevice>();

//    [Command("")]
    public void Count(string folderpath, string outfilepath)
    {
//== start
        logger.ZLogInformation($"==== tool {getMyFileVersion()} ====");
        if (!Directory.Exists(folderpath))
        {
            logger.ZLogError($"[NG] フォルダが見つかりません{folderpath}");
            return;
        }

        try
        {
            var excelpaths = Directory.GetFiles(folderpath);
            foreach (var excelpath in excelpaths)
            {
                List<MyDevicePort> mydeviceports = new List<MyDevicePort>();
                CreateFileToList(excelpath, mydeviceports);

                CountDevice(excelpath, mydeviceports);
                mydeviceports.Clear();
            }
        }
        catch (System.Exception ex)
        {
            logger.ZLogError($"ERROR {ex.ToString()}");
            throw;
        }

//== finish
        ExportFile(outfilepath);
        logger.ZLogInformation($"==== tool finish ====");
    }

    private void ExportFile(string outfilepath)
    {
        logger.ZLogInformation($"== 結果の出力 ==");
        string exportfilepath = getExportFileName(outfilepath);
        var sortKeys = MyUsedDevices.Keys.ToList();
        sortKeys.Sort();
        try
        {
            string deviceNameHeader = config.Value.DeviceNameHeader;
            using (StreamWriter file = new StreamWriter(exportfilepath, false, Encoding.GetEncoding("utf-8")))
            {
                file.WriteLine(deviceNameHeader);
                foreach (var key in sortKeys)
                {
                    var line = MyUsedDevices[key].siteNumberName + "," + MyUsedDevices[key].targetFileName + "," + MyUsedDevices[key].router + "," + MyUsedDevices[key].floorSw + "," + MyUsedDevices[key].poeSw + "," + MyUsedDevices[key].ap + "," + MyUsedDevices[key].lan + "," + MyUsedDevices[key].mc + "," + MyUsedDevices[key].floor;
                    file.WriteLine(line);
                    logger.ZLogInformation($"SiteName:{MyUsedDevices[key].siteNumberName},fileName:{MyUsedDevices[key].targetFileName},router:{MyUsedDevices[key].router},floorSw:{MyUsedDevices[key].floorSw},poeSw:{MyUsedDevices[key].poeSw},ap:{MyUsedDevices[key].ap},lan:{MyUsedDevices[key].lan},mc:{MyUsedDevices[key].mc},floor:{MyUsedDevices[key].floor}");
                }
            }
        }
        catch (System.Exception)
        {
            
            throw;
        }
    }

    private void CountDevice(string excelpath, List<MyDevicePort> mydeviceports)
    {
        logger.ZLogInformation($"== start Deviceのカウント ==");
        bool isError = false;

        string sieName = getSiteNameString(excelpath);
        var tmpDevice = new MyUsedDevice();
        tmpDevice.siteNumberName = sieName;
        tmpDevice.targetFileName = Path.GetFileName(excelpath);

        string deviceNameToRouter = config.Value.DeviceNameToRouter;
        List<string> listRouterName = deviceNameToRouter.Split(',').ToList<string>();
        string deviceNameToFloorSw = config.Value.DeviceNameToFloorSw;
        List<string> listFloorSwName = deviceNameToFloorSw.Split(',').ToList<string>();
        string deviceNameToPoeSw = config.Value.DeviceNameToPoeSw;
        List<string> listPoeSwName = deviceNameToPoeSw.Split(',').ToList<string>();
        string deviceNameToAp = config.Value.DeviceNameToAp;
        List<string> listApName = deviceNameToAp.Split(',').ToList<string>();
        string deviceNameToRosette = config.Value.DeviceNameToRosette;
        List<string> listLanName = deviceNameToRosette.Split(',').ToList<string>();
        string deviceNameToMc = config.Value.DeviceNameToMc;
        List<string> listMcName = deviceNameToMc.Split(',').ToList<string>();

        string wordConnect = config.Value.WordConnect;
        List<string> tmpRouter = new List<string>();
        List<string> tmpFloorSw = new List<string>();
        List<string> tmpPoeSw = new List<string>();
        List<string> tmpAp = new List<string>();
        List<string> tmpLan = new List<string>();
        List<string> tmpMc = new List<string>();
        List<string> tmpFloor = new List<string>();
        foreach (var device in mydeviceports)
        {
            if (device.fromConnect == wordConnect)
            {
                bool bMatch = false;
                // router
                if (isDevice(device.fromDeviceName, listRouterName))
                {
                    if (!tmpRouter.Contains(device.fromHostName))
                    {
                        bMatch = true;
                        tmpRouter.Add(device.fromHostName);
                        // floor
                        checkAndAddFloor(device.fromFloorName, tmpFloor);
                    }
                }
                // floorSw
                if (isDevice(device.fromDeviceName, listFloorSwName))
                {
                    if (!tmpFloorSw.Contains(device.fromHostName))
                    {
                        bMatch = true;
                        tmpFloorSw.Add(device.fromHostName);
                        // floor
                        checkAndAddFloor(device.fromFloorName, tmpFloor);
                    }
                }
                // poeSw
                if (isDevice(device.fromDeviceName, listPoeSwName))
                {
                    if (!tmpPoeSw.Contains(device.fromHostName))
                    {
                        bMatch = true;
                        tmpPoeSw.Add(device.fromHostName);
                        // floor
                        checkAndAddFloor(device.fromFloorName, tmpFloor);
                    }
                }
                // ap
                if (isDevice(device.toDeviceName, listApName))
                {
                    if (!tmpAp.Contains(device.toHostName))
                    {
                        bMatch = true;
                        tmpAp.Add(device.toHostName);
                        // floor
                        checkAndAddFloor(device.toFloorName, tmpFloor);
                    }
                }
                // lan == rosette
                if (isDevice(device.toDeviceName, listLanName))
                {
                    if (!tmpLan.Contains(device.toHostName))
                    {
                        bMatch = true;
                        tmpLan.Add(device.toHostName);
                        // floor
                        checkAndAddFloor(device.toFloorName, tmpFloor);
                    }
                }
                // mc
                if (isDevice(device.fromDeviceName, listMcName))
                {
                    if (!tmpMc.Contains(device.fromHostName))
                    {
                        bMatch = true;
                        tmpMc.Add(device.fromHostName);
                        // floor
                        checkAndAddFloor(device.fromFloorName, tmpFloor);
                    }
                }
                if (bMatch == false)
                {
                    logger.ZLogTrace($"[CountDevice] 不一致 ケーブルID:{device.fromCableID} From側デバイス名:{device.fromDeviceName} To側デバイス名:{device.toDeviceName}");
                }
            }
        }

        tmpDevice.router = tmpRouter.Count;
        tmpDevice.floorSw = tmpFloorSw.Count;
        tmpDevice.poeSw = tmpPoeSw.Count;
        tmpDevice.ap = tmpAp.Count;
        tmpDevice.lan = tmpLan.Count;
        tmpDevice.mc = tmpMc.Count;
        tmpDevice.floor = tmpFloor.Count;
        MyUsedDevices.Add(sieName, tmpDevice);

        if (isError)
        {
            isAllPass = false;
            logger.ZLogInformation($"[NG] Deviceのカウントでエラーが発生しました");
        }
        else
        {
            logger.ZLogInformation($"[OK] Deviceのカウントが正常に終了しました");
        }
        logger.ZLogInformation($"== end Deviceのカウント ==");
    }

    private void checkAndAddFloor(string floorName, List<string> tmpFloor)
    {
        if (!tmpFloor.Contains(floorName))
        {
            tmpFloor.Add(floorName);
        }
    }

    private bool isDevice(string device, List<string> listDevice)
    {
        return listDevice.Contains(device);
    }

    private string getExportFileName(string outfilepath)
    {
        string exportFolderPath = outfilepath;
        string exportFilename = DateTime.Now.ToString("yyyyMMdd")+".txt";
        return Path.Join(exportFolderPath, exportFilename);
    }

    private string getSiteNameString(string excelpath)
    {
        string fileNamePrifex = config.Value.FileNamePrifex;
        string fileNameWord = config.Value.FileNameWord;
        string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(excelpath);
        string replacePrifex = fileNameWithoutExtension.Replace(fileNamePrifex, "");
        int position1 = replacePrifex.IndexOf(fileNameWord);
        int position2 = replacePrifex.IndexOf(fileNameWord, position1+1);
        if (position2 < 0)
        {
            return replacePrifex;
        }
        string substring = replacePrifex.Substring(0, position2);
        return substring;
    }

    private void CreateFileToList(string excelpath, List<MyDevicePort> mydeviceports)
    {
        try
        {
            using FileStream fs = new FileStream(excelpath, FileMode.Open, FileAccess.Read, FileShare.Read);
            using XLWorkbook xlWorkbook = new XLWorkbook(fs);
            IXLWorksheets sheets = xlWorkbook.Worksheets;

//== init
            int deviceFromCableIdColumn = config.Value.DeviceFromCableIdColumn;
            int deviceFromKeyPortNameColumn = config.Value.DeviceFromKeyPortNameColumn;
            int deviceFromConnectColumn = config.Value.DeviceFromConnectColumn;
            int deviceFromFloorNameColumn = config.Value.DeviceFromFloorNameColumn;
            int deviceFromDeviceNameColumn = config.Value.DeviceFromDeviceNameColumn;
            int deviceFromDeviceNumberColumn = config.Value.DeviceFromDeviceNumberColumn;
            int deviceFromHostNameColumn = config.Value.DeviceFromHostNameColumn;
            int deviceFromModelNameColumn = config.Value.DeviceFromModelNameColumn;
            int deviceFromPortNameColumn = config.Value.DeviceFromPortNameColumn;
            int deviceFromConnectorNameColumn = config.Value.DeviceFromConnectorNameColumn;
            int deviceToFloorNameColumn = config.Value.DeviceToFloorNameColumn;
            int deviceToDeviceNameColumn = config.Value.DeviceToDeviceNameColumn;
            int deviceToDeviceNumberColumn = config.Value.DeviceToDeviceNumberColumn;
            int deviceToModelNameColumn = config.Value.DeviceToModelNameColumn;
            int deviceToHostNameColumn = config.Value.DeviceToHostNameColumn;
            int deviceToPortNameColumn = config.Value.DeviceToPortNameColumn;
            string wordConnect = config.Value.WordConnect;
            string wordDisconnect = config.Value.WordDisconnect;
            string ignoreDeviceNameToHostNameLength = config.Value.IgnoreDeviceNameToHostNameLength;
            string ignoreDeviceNameToHostNamePrefix = config.Value.IgnoreDeviceNameToHostNamePrefix;
            string ignoreDeviceNameToConnectXConnect = config.Value.IgnoreDeviceNameToConnectXConnect;
            string ignoreConnectorNameToAll = config.Value.IgnoreConnectorNameToAll;
            string wordDeviceToHostNameList = config.Value.WordDeviceToHostNameList;
            string deviceNameToRosette = config.Value.DeviceNameToRosette;

            logger.ZLogInformation($"== パラメーター ==");
            logger.ZLogInformation($"Checkファイル名:{excelpath}");

            foreach (IXLWorksheet? sheet in sheets)
            {
                if (sheet != null)
                {
                    int lastUsedRowNumber = sheet.LastRowUsed() == null ? 0 : sheet.LastRowUsed().RowNumber();
                    int lastUsedColumNumber = sheet.LastColumnUsed() == null ? 0 : sheet.LastColumnUsed().ColumnNumber();

//                    logger.ZLogInformation($"シート名:{sheet.Name}, 最後の行:{lastUsedRowNumber}, 最後の列:{lastUsedColumNumber}");

                    for (int r = 1; r < lastUsedRowNumber + 1; r++)
                    {
                        IXLCell cellConnect = sheet.Cell(r, deviceFromConnectColumn);
                        IXLCell cellCableID = sheet.Cell(r, deviceFromCableIdColumn);
                        if (cellConnect.IsEmpty() == true)
                        {
                            // nothing
                        }
                        else
                        {
                            if (cellConnect.Value.GetText() == wordConnect || cellConnect.Value.GetText() == wordDisconnect)
                            {
                                MyDevicePort tmpDevicePort = new MyDevicePort();
                                tmpDevicePort.fromConnect = cellConnect.Value.GetText();
                                int id = -1;
                                switch (cellCableID.DataType)
                                {
                                    case XLDataType.Number:
                                        id = cellCableID.GetValue<int>();
                                        break;
                                    case XLDataType.Text:
                                        try
                                        {
                                            id = int.Parse(cellCableID.GetValue<string>());
                                        }
                                        catch (System.FormatException)
                                        {
                                            isAllPass = false;
                                            logger.ZLogError($"[NG]ケーブルID is Error ( Text-> Int) at sheet:{sheet.Name} row:{r}");
                                            continue;
                                        }
                                        catch (System.Exception)
                                        {
                                            throw;
                                        }
                                        break;
                                    default:
                                        isAllPass = false;
                                        logger.ZLogError($"[NG]ケーブルID is NOT type ( Number | Text ) at sheet:{sheet.Name} row:{r}");
                                        continue;
                                }
                                tmpDevicePort.fromCableID = id;
                                tmpDevicePort.fromKeyPortName = sheet.Cell(r, deviceFromKeyPortNameColumn).Value.ToString();
                                tmpDevicePort.fromFloorName = sheet.Cell(r, deviceFromFloorNameColumn).Value.ToString();
                                tmpDevicePort.fromDeviceName = sheet.Cell(r, deviceFromDeviceNameColumn).Value.ToString();
                                tmpDevicePort.fromDeviceNumber = sheet.Cell(r, deviceFromDeviceNumberColumn).Value.ToString();
                                tmpDevicePort.fromHostName = sheet.Cell(r, deviceFromHostNameColumn).Value.ToString();
                                tmpDevicePort.fromModelName = sheet.Cell(r, deviceFromModelNameColumn).Value.ToString();
                                tmpDevicePort.fromPortName = sheet.Cell(r, deviceFromPortNameColumn).Value.ToString();
                                tmpDevicePort.fromConnectorName = sheet.Cell(r, deviceFromConnectorNameColumn).Value.ToString();
                                tmpDevicePort.toFloorName = sheet.Cell(r, deviceToFloorNameColumn).Value.ToString();
                                tmpDevicePort.toDeviceName = sheet.Cell(r, deviceToDeviceNameColumn).Value.ToString();
                                tmpDevicePort.toDeviceNumber = sheet.Cell(r, deviceToDeviceNumberColumn).Value.ToString();
                                tmpDevicePort.toModelName = sheet.Cell(r, deviceToModelNameColumn).Value.ToString();
                                tmpDevicePort.toHostName = sheet.Cell(r, deviceToHostNameColumn).Value.ToString();
                                tmpDevicePort.toPortName = sheet.Cell(r, deviceToPortNameColumn).Value.ToString();
                                mydeviceports.Add(tmpDevicePort);
                            }
                        }
                    }
                }
            }

//== print
            printMyDevicePorts(mydeviceports);

        }
        catch (IOException ie)
        {
            logger.ZLogError($"[ERROR] Excelファイルの読み取りでエラーが発生しました。Excelで対象ファイルを開いていませんか？ 詳細:({ie.Message})");
            return;
        }
        catch (System.Exception)
        {
            throw;
        }

    }

    private void printMyDevicePorts(List<MyDevicePort> mydeviceports)
    {
        logger.ZLogTrace($"== start print ==");
        foreach (var device in mydeviceports)
        {
            logger.ZLogTrace($"CableID:{device.fromCableID},connect:{device.fromConnect},(from) Device:{device.fromDeviceName},Host:{device.fromHostName},Model:{device.fromModelName},Port:{device.fromPortName},(to) Device:{device.toDeviceName},Host:{device.toHostName},Model:{device.toModelName},Port:{device.toPortName}");
        }
        logger.ZLogTrace($"== end print ==");
    }

    private string getMyFileVersion()
    {
        System.Diagnostics.FileVersionInfo ver = System.Diagnostics.FileVersionInfo.GetVersionInfo(System.Reflection.Assembly.GetExecutingAssembly().Location);
        return ver.InternalName + "(" + ver.FileVersion + ")";
    }
}

//==
public class MyConfig
{
    public string Header {get; set;} = "";

    public string ModelAndPortName {get; set;} = "";
    public string IgnoreModelName {get; set;} = "";
    public int DeviceFromCableIdColumn {get; set;} = -1;
    public int DeviceFromKeyPortNameColumn {get; set;} = -1;
    public int DeviceFromConnectColumn {get; set;} = -1;
    public int DeviceFromFloorNameColumn {get; set;} = -1;
    public int DeviceFromDeviceNameColumn {get; set;} = -1;
    public int DeviceFromDeviceNumberColumn {get; set;} = -1;
    public int DeviceFromModelNameColumn {get; set;} = -1;
    public int DeviceFromHostNameColumn {get; set;} = -1;
    public int DeviceFromPortNameColumn {get; set;} = -1;
    public int DeviceFromConnectorNameColumn {get; set;} = -1;
    public int DeviceToFloorNameColumn {get; set;} = -1;
    public int DeviceToDeviceNameColumn {get; set;} = -1;
    public int DeviceToDeviceNumberColumn {get; set;} = -1;
    public int DeviceToModelNameColumn {get; set;} = -1;
    public int DeviceToHostNameColumn {get; set;} = -1;
    public int DeviceToPortNameColumn {get; set;} = -1;
    public string WordConnect {get; set;} = "";
    public string WordDisconnect {get; set;} = "";
    public string IgnoreDeviceNameToHostNameLength {get; set;} = "";
    public string IgnoreDeviceNameToHostNamePrefix {get; set;} = "";
    public string IgnoreDeviceNameToConnectXConnect {get; set;} = "";
    public string IgnoreConnectorNameToAll {get; set;} = "";
    public string WordDeviceToHostNameList {get; set;} = "";
    public string DeviceNameHeader {get; set;} = "";
    public string DeviceNameToRouter {get; set;} = "";
    public string DeviceNameToFloorSw {get; set;} = "";
    public string DeviceNameToPoeSw {get; set;} = "";
    public string DeviceNameToAp {get; set;} = "";
    public string DeviceNameToRosette {get; set;} = "";
    public string DeviceNameToMc {get; set;} = "";
    public string FileNamePrifex {get; set;} = "";
    public string FileNameWord {get; set;} = "";
}

public class MyDevicePort
{
    public int fromCableID = -1;
    public string fromConnect = "";

    public string fromKeyPortName = "";
    public string fromFloorName = "";
    public string fromDeviceName = "";
    public string fromDeviceNumber = "";
    public string fromModelName = "";
    public string fromHostName = "";
    public string fromPortName = "";
    public string fromConnectorName = "";

    public string toFloorName = "";
    public string toDeviceName = "";
    public string toDeviceNumber = "";
    public string toModelName = "";
    public string toHostName = "";
    public string toPortName = "";
}

public class MyUsedDevice
{
    public string siteNumberName = "";
    public string targetFileName = "";
    public int router = -1;
    public int floorSw = -1;
    public int poeSw = -1;
    public int ap = -1;
    public int lan = -1;
    public int mc = -1;
    public int floor = -1;
}
