using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Spire.Doc;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

namespace Tool.Doc
{
    public class WordTool
    {
        private readonly string _rootPath;
        private string _samplePath;
        private byte[] _sampleByteArr;
        private string _tempPath;
        private int _randomIndex = 0;
        private string _randomStr
        {
            get
            {
                return $"{DateTime.Now.ToString("yyyy_MM_dd_hh_mm_ss")}_{++_randomIndex}";
            }
        }
        public WordTool(string rootPath)
        {
            _rootPath = rootPath;
        }
        public WordTool LoadFile(byte[] _fileSrc)
        {
            this._sampleByteArr = _fileSrc;
            this._tempPath = Path.Combine(_rootPath, "tempData", $"{_randomStr}_file");
            return this;
        }
        public byte[] FindAndReplace(Dictionary<string, object> keyValueDict)
        {
            try
            {
                //複製
                File.WriteAllBytes(this._tempPath, this._sampleByteArr);
                //處理
                using (WordprocessingDocument doc = WordprocessingDocument.Open(this._tempPath, true))
                {
                    var elem = doc.MainDocumentPart.Document.Body;
                    var tempPool = new List<Run>();
                    var matchText = string.Empty;
                    var highlightRuns = elem.Descendants<Run>().Where(x => x.RunProperties?.Elements<Highlight>().Any() ?? false).ToList();
                    //找尋
                    foreach (var hlRun in highlightRuns)
                    {
                        var text = hlRun.InnerText;
                        if (text.StartsWith("{"))
                        {
                            tempPool = new List<Run>() { hlRun };
                            matchText = text;
                        }
                        else
                        {
                            matchText = matchText + text;
                            tempPool.Add(hlRun);
                        }
                        if (text.EndsWith("}"))
                        {
                            var m = Regex.Match(matchText, @"\{\$(?<n>\w+)\$\}");
                            if (m.Success && keyValueDict.ContainsKey(m.Groups["n"].Value))
                            {
                                var firstRun = tempPool.First();
                                firstRun.RemoveAllChildren<Text>();
                                firstRun.RunProperties.RemoveAllChildren<Highlight>();
                                var valObject = keyValueDict[m.Groups["n"].Value];
                                //分類型處理
                                if (valObject is DocumentFormat.OpenXml.Drawing.Table)
                                {
                                    //TODO
                                }
                                else
                                {
                                    var firstLine = true;
                                    foreach (var line in Regex.Split(valObject.ToString(), @"\\n"))
                                    {
                                        if (firstLine) firstLine = false;
                                        else firstRun.Append(new DocumentFormat.OpenXml.Drawing.Break());
                                        firstRun.Append(new Text(line));
                                    }
                                    tempPool.Skip(1).ToList().ForEach(o => o.Remove());
                                }
                            }
                        }
                    }
                }
                //輸出
                return File.ReadAllBytes(this._tempPath);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                File.Delete(this._tempPath);
            }
        }

        public byte[] ToPDF()
        {
            try
            {
                //複製
                File.WriteAllBytes(this._tempPath, this._sampleByteArr);
                var document = new Spire.Doc.Document();
                document.LoadFromFile(this._tempPath);
                var byteArr = new byte[0];
                using (var pdfMemoryStream = new MemoryStream())
                {
                    document.SaveToFile(pdfMemoryStream, FileFormat.PDF);
                    byteArr = pdfMemoryStream.ToArray();
                }
                return byteArr;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                File.Delete(this._tempPath);
            }
        }
    }
}
