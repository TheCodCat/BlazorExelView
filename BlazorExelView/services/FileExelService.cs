using BlazorExelView.model;
using OfficeOpenXml;

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
                            if(string.Equals(organization, "ГБПОУ СО «Самарский государственный колледж»")){
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
        public async Task<Dictionary<string,List<Worker>>> LoadFilePages(string urlstr)
        {
            FileInfo fileInfo = new FileInfo(urlstr);
            Dictionary<string, List<Worker>> pagesWorker = new Dictionary<string, List<Worker>>();
            if (fileInfo.Exists)
            {
                using (ExcelPackage package = new ExcelPackage(fileInfo))
                {
                    foreach (var excelWorksheet in package.Workbook.Worksheets)
                    {
                        List<Worker> workers = new List<Worker>();
                        var totalCols = excelWorksheet.Dimension.End.Column;
                        var totalrow = 100;
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
                                    worker.Description = workWorker.ToString();

                                    var valueWorker = excelWorksheet.GetValue(j, i);
                                    worker.State = valueWorker != null ? valueWorker.ToString()  : null;
                                    worker.Init();

                                    workers.Add(worker);
                                }
                            }
                        }
                        var sorted = await SortWorkers(workers);
                        pagesWorker.Add(excelWorksheet.Name,sorted);
                    }
                }

            }
            return pagesWorker;
        }

        private async Task<List<Worker>> SortWorkers(List<Worker> workers)
        {
            var sort = workers.GroupBy(x => x.ID);
            List<Worker> sortedWorker = new List<Worker>();
            foreach (var item in sort)
            {
                int index = 0;
                foreach (var work in item)
                {
                    index++;

                    work.IDView = $"{work.ID}.{index}";
                    sortedWorker.Add(work);
                }
            }
            return sortedWorker;
        }
    }
}
