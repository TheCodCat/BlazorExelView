using BlazorExelView.model;
using OfficeOpenXml;
using System.Text.Json;

namespace BlazorExelView.services
{
    public class FileExelService
    {
        public async Task<List<Worker>> LoadFile(string urlstr)
        {
            List<Worker> workers = new List<Worker>();
            FileInfo fileInfo = new FileInfo(urlstr);

            if (fileInfo.Exists)
            {
                using (ExcelPackage package = new ExcelPackage(fileInfo))
                {
                    ExcelWorksheet excelWorksheet = package.Workbook.Worksheets.First();

                    var totalCols = excelWorksheet.Dimension.End.Column;
                    var totalrow = 60;
                    for (int i = 3; i < totalCols; i++)
                    {
                        var workername = excelWorksheet.Cells[1, i].Value;
                        var workWorker = excelWorksheet.Cells[2, i].Value;
                        if (workername == null) continue;

                        Worker worker = new Worker(workername?.ToString() ?? "непонятно", i - 2);
                        worker.Description = workWorker.ToString();
                        for (int j = 2; j < totalrow; j++)
                        {
                            var organization = excelWorksheet.GetValue(j, 2);
                            if (string.Equals(organization, "ГБПОУ СО «Самарский государственный колледж»"))
                            {
                                var valueWorker = excelWorksheet.GetValue(j, i);
                                worker.State = valueWorker != null ? valueWorker.ToString() : "Не в работе";
                            }
                        }
                        workers.Add(worker);
                    }
                }

            }
            return workers;
        }
        public async Task<Dictionary<string, List<Worker>>> LoadFilePages(string urlstr)
        {
            try
            {
                Dictionary<string, List<Worker>> pagesWorker = new Dictionary<string, List<Worker>>();
                using (var httpclient = new HttpClient())
                {
                    string url = $"https://cloud-api.yandex.net/v1/disk/public/resources/download?public_key={Uri.EscapeDataString(urlstr)}";
                    var json = await httpclient.GetStringAsync(url);
                    using JsonDocument doc = JsonDocument.Parse(json);
                    JsonElement root = doc.RootElement;
                    string downloadUrl = root.GetProperty("href").GetString();

                    byte[] buffer = await httpclient.GetByteArrayAsync(downloadUrl);

                    using (var stream = new MemoryStream(buffer))
                    {
                        using (ExcelPackage package = new ExcelPackage(stream))
                        {
                            foreach (var excelWorksheet in package.Workbook.Worksheets)
                            {
                                List<Worker> workers = new List<Worker>();

                                var totalrow = 3;
                                while (true)
                                {
                                    if (excelWorksheet.Cells[totalrow, 2].Value is string)
                                    {
                                        totalrow++;
                                        //Console.WriteLine($"{totalrow} строка");
                                    }
                                    else
                                        break;
                                }
                                var totalCols = 3;
                                while (true)
                                {
                                    if (excelWorksheet.Cells[1, totalCols].Value is string)
                                    {
                                        totalCols++;
                                        //Console.WriteLine($"{totalCols} колонна");
                                    }
                                    else
                                        break;
                                }


                                for (int i = 3; i < totalCols; i++)
                                {
                                    var workername = excelWorksheet.Cells[1, i].Value;
                                    var workWorker = excelWorksheet.Cells[2, i].Value;
                                    if (workername == null) continue;


                                    for (int j = 2; j < totalrow; j++)
                                    {
                                        var organization = excelWorksheet.GetValue(j, 2);
                                        if (string.Equals(organization, "ГАПОУ СО «Самарский государственный колледж»"))
                                        {
                                            Worker worker = new Worker(workername?.ToString() ?? "непонятно", i - 2);
                                            worker.Description = workWorker is string ? workWorker.ToString() : string.Empty;

                                            var valueWorker = excelWorksheet.GetValue(j, i);
                                            worker.State = valueWorker is string ? valueWorker.ToString() : string.Empty;
                                            worker.Init();

                                            workers.Add(worker);
                                        }
                                    }
                                }
                                var sorted = await SortWorkers(workers);
                                pagesWorker.Add(excelWorksheet.Name, sorted);
                            }
                        }
                    }
                }
                return pagesWorker;
            }
            catch (Exception ex)
            {
                return new Dictionary<string, List<Worker>>();
            }

        }
        private async Task<List<Worker>> SortWorkers(List<Worker> workers)
        {
            var sort = workers.GroupBy(x => x.ID);
            List<Worker> sortedWorker = new List<Worker>();
            foreach (var item in sort)
            {
                if (item.Count() > 1)
                {
                    int index = 0;
                    foreach (var work in item)
                    {
                        index++;

                        work.IDView = $"{work.ID}.{index}";
                        sortedWorker.Add(work);
                    }
                }
                else
                {
                    foreach (var work in item)
                    {
                        work.IDView = $"{work.ID}";
                        sortedWorker.Add(work);
                    }
                }
            }
            return sortedWorker;
        }
    }
}
