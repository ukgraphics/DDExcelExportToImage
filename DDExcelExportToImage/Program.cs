using GrapeCity.Documents.Excel;

namespace DDExcelExportToImage
{
    class Program
    {
        static void Main(string[] args)
        {
            // ライセンスを設定します
            //Workbook.SetLicenseKey("トライアル版もしくは製品版のライセンスキー");

            // 新規ワークブックの作成
            var workbook = new Workbook();

            // xlsxファイルを開く
            workbook.Open("Template_SalesTracker_report.xlsx");
            IWorksheet worksheet = workbook.Worksheets[0];

            // ワークシートを画像に出力
            worksheet.ToImage("worksheet.png");

            // セル範囲を画像に出力
            worksheet.Range["B15:K21"].ToImage("range.png");

            // シェイプを画像に出力
            worksheet.Shapes["TextBox"].ToImage("textbox.png");

            // バーチャートを画像に出力
            worksheet.Shapes["ProductIncomeChart"].ToImage("barchart.png");

            // パイチャートを画像に出力
            worksheet.Shapes["ProductIncomePctChart"].ToImage("piechart.png");

        }
    }
}
