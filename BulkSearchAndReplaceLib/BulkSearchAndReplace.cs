using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Excel;
using Microsoft.Office.Interop.Word;
using Novacode;

namespace BulkSearchAndReplaceLib
{
    public class BulkSearchAndReplace
    {
        private const String TmpDirname = "tmp";
        private const String Word2003Extension = ".doc";
        private const String Word2007Extension = ".docx";
        private const String Excel2003Extension = ".xls";
        private const String Excel2007Extension = ".xlsx";
        private const String DirectorySeperator = "\\";

        public const String FileFilterTxt = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
        public const String FileFilterExcel = "exel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";

        private static BulkSearchAndReplace _instance;
        private String _configFilePath;
        private String _destinationDirectoryPath;
        private String _sourceDirectoryPath;
        private String _excelFilePath;
        private bool _isRunning;

        private Dictionary<string, string> _replaceItemList;
        private readonly Dictionary<String, String> _directories = new Dictionary<String, String>();
        private readonly List<String> _sourceFiles = new List<String>();
        private readonly bool _useDoc2003 = true;

        private BulkSearchAndReplace()
        {

            // nothing to do
        }


        public string ConfigFilePath
        {
            set { _configFilePath = value; }
        }


        public string SourceDirectorPath
        {
            set { _sourceDirectoryPath = value; }

        }


        public string DestinationDirectoryPath
        {
            set { _destinationDirectoryPath = value; }
        }


        public string ExcelFilePath
        {
            set { _excelFilePath = value; }
        }


        public static BulkSearchAndReplace GetInstance()
        {
            return _instance = _instance ?? new BulkSearchAndReplace();
        }

        /// <summary>
        /// Initial worker
        /// </summary>
        /// <returns>string</returns>
        public string Run()
        {
            _isRunning = true;

            if (_configFilePath != null)
            {
                if (ExtractConfig() == false)
                {
                    CleanUp();
                    return "Please use a valid config format.";
                }
            }

            if (_sourceDirectoryPath == null)
            {
                CleanUp();
                return "Please set a valid source directory.";
            }

            if (_destinationDirectoryPath == null)
            {
                CleanUp();
                return "Please set a valid destination directory.";
            }

            if (_excelFilePath == null)
            {
                CleanUp();
                return "Please set a valid exel file.";
            }


            if (ReadExelDataToDictionary() != true)
            {
                CleanUp();
                return "Could not parse excel file.\nIs It maybe open in another programm?";
            }

            GetDirectoriesRecursive(_sourceDirectoryPath);

            if (GetFileNamesInDir() != true)
            {
                CleanUp();
                return "The source directory has no word documents.";
            }

            RepleaceInWord();
            CleanUp();

            return "Success!";
        }

        /// <summary>
        /// Call directory cleanup and reset run var
        /// </summary>
        private void CleanUp()
        {
            _isRunning = false;
            CleanUpTmpFolder();
        }

        /// <summary>
        /// Delete tmp files and directories
        /// </summary>
        private void CleanUpTmpFolder()
        {
            if (_directories.Any())
            {
                return;
            }

            foreach (var keyValuePair in _directories)
            {
                var subDirectories = Directory.GetDirectories(keyValuePair.Key);
                foreach (var subDirectory in subDirectories)
                {
                    if (subDirectory.Contains(TmpDirname))
                    {
                        var directoryInfo = new DirectoryInfo(subDirectory);
                        if (directoryInfo.FullName.Contains(TmpDirname))
                        {

                            foreach (var fileInfo in directoryInfo.GetFiles())
                            {
                                fileInfo.Delete();
                            }

                            directoryInfo.Delete();
                        }
                    }
                }

            }

        }

        /// <summary>
        /// Find all subdirectories
        /// </summary>
        /// <param name="sourceDirectorPath">String</param>
        private void GetDirectoriesRecursive(String sourceDirectorPath)
        {
            var directoryInfo = new DirectoryInfo(sourceDirectorPath);
            var subDirectories = directoryInfo.GetDirectories();

            foreach (var subDiretory in subDirectories)
            {
                var path = subDiretory.FullName;
                var relevantPathPart = path.Replace(_sourceDirectoryPath, "");
                var destinationPath = _destinationDirectoryPath + relevantPathPart;

                if (!path.Contains(TmpDirname))
                {
                    _directories.Add(path, destinationPath);
                    GetDirectoriesRecursive(subDiretory.FullName);
                }
            }
        }


        /// <summary>
        /// Get all relevant files from directories an subdirectories
        /// </summary>
        /// <returns>bool</returns>
        private bool GetFileNamesInDir()
        {
            var filter = "*" + Word2007Extension;
            if (_useDoc2003)
            {
                filter += ";*" + Word2003Extension;
            }

            var extensionSplit = filter.Split(';');
            foreach (var filterItem in extensionSplit)
            {
                foreach (var keyValuePair in _directories)
                {
                    var files = Directory.GetFiles(keyValuePair.Key, filterItem);
                    foreach (var file in files)
                    {
                        if (!_sourceFiles.Contains(file))
                        {
                            _sourceFiles.Add(file);
                        }
                    }
                }
            }

            if (_sourceFiles.Count > 0)
            {
                return true;
            }

            return false;
        }

        /// <summary>
        /// Use Microsoft.Office.Interop.Word to open word2003 and convert to word2007
        /// </summary>
        /// <param name="fileNameWithPath">String</param>
        /// <returns>String</returns>
        private String ConvertDoc2DocX(String fileNameWithPath)
        {
            var tmpFilenameWithPath = CreateTempDirAndCopyTmpFile(fileNameWithPath);
            if (tmpFilenameWithPath == null)
            {
                return null;
            }
            var wordApplication = new Application();
            wordApplication.Visible = false;

            var wordDocument = wordApplication.Documents.Open(
                tmpFilenameWithPath, false, false
                );
            wordDocument.Activate();

            object fileFormat = WdSaveFormat.wdFormatXMLDocument;
            var outputFileName = wordDocument.FullName.Replace(Word2003Extension, Word2007Extension);
            wordDocument.SaveAs2(outputFileName, fileFormat,
                CompatibilityMode: WdCompatibilityMode.wdWord2010);

            wordDocument.Close();
            wordApplication.Quit();

            return outputFileName;
        }

        /// <summary>
        /// Create tmp dir and copy current word2003 document, don't work with original
        /// </summary>
        /// <param name="fileNameWithPath">String</param>
        /// <returns>string</returns>
        private string CreateTempDirAndCopyTmpFile(String fileNameWithPath)
        {
            var path = Path.GetFullPath(fileNameWithPath);
            var fileName = Path.GetFileNameWithoutExtension(fileNameWithPath);
            var pathParts = path.Split(new[] { DirectorySeperator }, StringSplitOptions.None);
            if (pathParts.Length > 0)
            {
                path = path.Replace(pathParts.Last(), TmpDirname);

                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }
            }

            if (Directory.Exists(path))
            {
                var newFullFilePath = path + DirectorySeperator + fileName + Word2003Extension;
                File.Copy(fileNameWithPath, newFullFilePath, true);

                var newFileExists = File.Exists(newFullFilePath);
                if (newFileExists)
                {
                    return newFullFilePath;
                }
            }

            return null;
        }

        /// <summary>
        /// Replace current key value pair from excel, if it's a word2003 document call converter
        /// </summary>
        /// <returns>bool</returns>
        private bool RepleaceInWord()
        {
            foreach (var fileNameWithPath in _sourceFiles)
            {
                var fileNameWithPathTmp = fileNameWithPath;
                if (File.Exists(fileNameWithPath))
                {
                    var fileName = Path.GetFileName(fileNameWithPathTmp);
                    var extension = Path.GetExtension(fileNameWithPathTmp);

                    if (Word2003Extension.Equals(extension))
                    {
                        if (_useDoc2003)
                        {
                            fileNameWithPathTmp = ConvertDoc2DocX(fileNameWithPathTmp);
                            fileName = Path.GetFileName(fileNameWithPathTmp);
                        }
                        else
                        {
                            continue;
                        }
                    }

                    if (fileNameWithPathTmp != null)
                    {
                        var document = DocX.Load(fileNameWithPathTmp);
                        foreach (var entry in _replaceItemList)
                        {
                            document.ReplaceText(entry.Key, entry.Value);
                        }


                        SaveFile(document, fileNameWithPath, fileName);
                    }
                }
            }

            return true;
        }

        /// <summary>
        /// Save file in same subdirectory in destination directoy path, if directory not exists create it
        /// </summary>
        /// <param name="document">DocX</param>
        /// <param name="fileNameWithPath">String</param>
        /// <param name="fileName">String</param>
        private void SaveFile(DocX document, String fileNameWithPath, String fileName)
        {
            var path = Path.GetDirectoryName(fileNameWithPath);

            foreach (var item in _directories)
            {
                if (path.Equals(item.Key))
                {
                    if (!Directory.Exists(item.Value))
                    {
                        Directory.CreateDirectory(item.Value);
                    }

                    document.SaveAs(item.Value + DirectorySeperator + fileName);
                }
            }
        }

        /// <summary>
        /// Open excel file
        /// </summary>
        /// <returns>FileStream</returns>
        private FileStream OpenFile()
        {
            try
            {
                var stream = File.Open(_excelFilePath, FileMode.Open, FileAccess.Read);
                return stream;
            }
            catch (Exception e)
            {
                return null;
            }
        }

        /// <summary>
        /// Read Excel data in a Dictionary list
        /// </summary>
        /// <returns>bool</returns>
        private bool ReadExelDataToDictionary()
        {
            IExcelDataReader excelReader = null;
            var file = new FileInfo(_excelFilePath);
            var stream = OpenFile();
            if (stream == null)
            {
                return false;
            }
            if (file.Extension == Excel2003Extension)
            {
                excelReader = ExcelReaderFactory.CreateBinaryReader(stream);
            }
            else if (file.Extension == Excel2007Extension)
            {
                excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
            }

            if (excelReader == null)
            {
                return false;
            }

            excelReader.IsFirstRowAsColumnNames = true;
            _replaceItemList = new Dictionary<string, string>();
            while (excelReader.Read())
            {
                var key = excelReader.GetString(0);
                var value = excelReader.GetString(1);
                _replaceItemList.Add(key, value);
            }

            excelReader.Close();

            return true;
        }

        /// <summary>
        /// Extract config and save config data in class var
        /// </summary>
        /// <returns>bool</returns>
        public bool ExtractConfig()
        {
            var lines = File.ReadAllLines(_configFilePath);
            foreach (var line in lines)
            {
                var keyValueStrings = line.Split('=');
                try
                {
                    var key = keyValueStrings[0];
                    var value = keyValueStrings[1];

                    switch (key)
                    {
                        case "source":
                            _sourceDirectoryPath = value;
                            break;
                        case "destination":
                            _destinationDirectoryPath = value;
                            break;
                        case "excel":
                            _excelFilePath = value;
                            break;
                        default:
                            break;
                    }
                }
                catch (Exception)
                {
                    return false;
                }
            }
            return true;
        }
    }
}