
using System.Globalization;
using System.Text.RegularExpressions;
using Xceed.Document.NET;
using Xceed.Words.NET;
using System.Configuration;


namespace ReportGeneratorPro
{
    public class ReportData
    {
        public string Number { get; set; } // 编号
        public string UsingUnit { get; set; } // 使用单位
        public string EntrustedUnit { get; set; } // 委托单位
        public string ElevatorRegistrationCode { get; set; } // 电梯注册代码
        public string ElevatorFactoryCode { get; set; } // 电梯出厂编号
        public string GovernorModel { get; set; } // 限速器型号
        public string GovernorFactoryCode { get; set; } // 限速器出厂编号
        public string RatedSpeed { get; set; } // 额定速度
        public string EquipmentForm { get; set; } // 设备形式

        // 新增字段
        public string DownwardElectricalAverageSpeed { get; set; } // 下行电气动作速度 - 平均值
        public string DownwardElectricalEvaluation { get; set; } // 下行电气动作速度 - 评价
        public string DownwardMechanicalAverageSpeed { get; set; } // 下行机械动作速度 - 平均值
        public string DownwardMechanicalEvaluation { get; set; } // 下行机械动作速度 - 评价
        public string UpwardElectricalAverageSpeed { get; set; } // 上行电气动作速度 - 平均值
        public string UpwardElectricalEvaluation { get; set; } // 上行电气动作速度 - 评价
        public string UpwardMechanicalAverageSpeed { get; set; } // 上行机械动作速度 - 平均值
        public string UpwardMechanicalEvaluation { get; set; } // 上行机械动作速度 - 评价
        public string ElectricalSafetyDevice1 { get; set; } // 电气安全装置 - 第一个文本框
        public string ElectricalSafetyDevice2 { get; set; } // 电气安全装置 - 第二个文本框

        // 新增字段
        public string Notes { get; set; } // 备注
        public string InspectionDate { get; set; } // 校验日期
        public string NextInspectionDate { get; set; } // 下次校验日期
    }

    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            try
            {
                // 阶段1：数据采集
                
                var sourcePaths = SelectMultipleFiles("选择限速器测试记录文件", "Word文档|*.docx");
                var templatePath = GetTemplatePath();

                // 阶段2：数据处理和报告生成
                foreach (var sourcePath in sourcePaths)
                {
                    var reportData = DataProcessingPhase(sourcePath);
                    ReportGenerationPhase(templatePath, reportData, sourcePath);
                }
            }
            catch (OperationCanceledException)
            {
                Console.WriteLine("操作已取消");
            }
            catch (Exception ex)
            {
                ShowColoredMessage($"严重错误：{ex.Message}", ConsoleColor.Red);
            }
        }

        // 获取报告模板路径
        static string GetTemplatePath()
        {
            string templatePath = ConfigurationManager.AppSettings["TemplatePath"];

            if (string.IsNullOrEmpty(templatePath))
            {
                templatePath = SelectFile("选择报告模板", "Word模板|*.docx");
                if (string.IsNullOrEmpty(templatePath)) throw new OperationCanceledException();

                // 保存模板路径到配置文件
                Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
                config.AppSettings.Settings.Add("TemplatePath", templatePath);
                config.Save(ConfigurationSaveMode.Modified);
                ConfigurationManager.RefreshSection("appSettings");
            }

            return templatePath;
        }

        // 文件选择阶段
        static string SelectFile(string title, string filter)
        {
            using var dialog = new OpenFileDialog
            {
                Title = title,
                Filter = filter,
                CheckFileExists = true,
                Multiselect = false
            };
            return dialog.ShowDialog() == DialogResult.OK ? dialog.FileName : null;
        }

        // 多文件选择
        static List<string> SelectMultipleFiles(string title, string filter)
        {
            using var dialog = new OpenFileDialog
            {
                Title = title,
                Filter = filter,
                CheckFileExists = true,
                Multiselect = true
            };
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                return dialog.FileNames.ToList();
            }
            return new List<string>();
        }

        // 数据处理阶段
        static ReportData DataProcessingPhase(string filePath)
        {
            using var doc = DocX.Load(filePath);

            var data = new ReportData();

            // 提取编号 RTE-JX 后面的内容
            data.Number = ExtractNumber(doc);

            // 提取表格中的数据
            var table = doc.Tables.LastOrDefault(t => t.Rows.Count >= 6); // 假设表格是文档中的第一个表格
            if (table != null)
            {
                // 基本信息
                data.UsingUnit = GetCellText(table, 1, 2); // 使用单位（第1行第2列）
                data.EntrustedUnit = GetCellText(table, 2, 2); // 委托单位（第2行第2列）
                data.ElevatorRegistrationCode = GetCellText(table, 3, 2); // 电梯注册代码（第3行第2列）
                data.ElevatorFactoryCode = GetCellText(table, 4, 2); // 电梯出厂编号（第4行第2列）
                data.GovernorModel = GetCellText(table, 5, 3); // 限速器型号（第5行第3列）
                data.GovernorFactoryCode = GetCellText(table, 5, 5); // 限速器出厂编号（第5行第5列）
                data.RatedSpeed = GetCellText(table, 6, 3); // 额定速度（第6行第3列）
                data.EquipmentForm = GetCellText(table, 6, 5) == "☑  单向 ☐  双向" ? "单向" : "双向"; // 设备形式（第6行第5列）

                // 新增内容
                data.DownwardElectricalAverageSpeed = ConverSpeedText(GetCellText(table, 10, 5)); // 下行电气动作速度 - 平均值
                data.DownwardElectricalEvaluation = ConvertEvaluation(GetCellText(table, 10, 6)); // 下行电气动作速度 - 评价

                data.DownwardMechanicalAverageSpeed = ConverSpeedText(GetCellText(table, 11, 5)); // 下行机械动作速度 - 平均值
                data.DownwardMechanicalEvaluation = ConvertEvaluation(GetCellText(table, 11, 6)); // 下行机械动作速度 - 评价

                data.UpwardElectricalAverageSpeed = ConverSpeedText(GetCellText(table, 12, 5)); // 上行电气动作速度 - 平均值
                data.UpwardElectricalEvaluation = ConvertEvaluation(GetCellText(table, 12, 6)); // 上行电气动作速度 - 评价

                data.UpwardMechanicalAverageSpeed = ConverSpeedText(GetCellText(table, 13, 5)); // 上行机械动作速度 - 平均值
                data.UpwardMechanicalEvaluation = ConvertEvaluation(GetCellText(table, 13, 6)); // 上行机械动作速度 - 评价

                data.ElectricalSafetyDevice1 = ConvertSafetyDevice(GetCellText(table, 14, 2)); // 电气安全装置 - 第一个文本框
                data.ElectricalSafetyDevice2 = ConvertEvaluation(GetCellText(table, 14, 3)); // 电气安全装置 - 第二个文本框

                // 新增内容
                data.Notes = GetCellText(table, 18, 2);
                data.InspectionDate = GetCellText(table, 19, 4);
                data.NextInspectionDate = GetCellText(table, 20, 4);
            }

            return data;
        }

        // 报告生成阶段
        static void ReportGenerationPhase(string templatePath, ReportData data, string sourcePath)
        {
            using var doc = DocX.Load(templatePath);

            // 查找包含 “编号：BTE-JX” 的段落
            var targetParagraph = doc.Paragraphs.FirstOrDefault(p => p.Text.Contains("编号：BTE-JX"));

            if (targetParagraph != null)
            {
                // 添加新的文本到段落
                targetParagraph.Append(data.Number);
            }

            // 找到表格并写入数据
            var resultTable = doc.Tables.LastOrDefault(t => t.Rows.Count >= 6); // 假设结果表格是第一个表格
            if (resultTable != null)
            {
                // 基本信息
                SetTableCellText(resultTable.Rows[0].Cells[1], data.UsingUnit ?? "—");
                SetTableCellText(resultTable.Rows[1].Cells[1], data.EntrustedUnit ?? "—");
                SetTableCellText(resultTable.Rows[2].Cells[1], data.ElevatorRegistrationCode ?? "—");
                SetTableCellText(resultTable.Rows[3].Cells[1], data.ElevatorFactoryCode ?? "—");

                SetTableCellText(resultTable.Rows[4].Cells[2], data.GovernorModel ?? "—");
                SetTableCellText(resultTable.Rows[4].Cells[4], data.GovernorFactoryCode ?? "—");
                SetTableCellText(resultTable.Rows[5].Cells[2], data.RatedSpeed ?? "—");
                SetTableCellText(resultTable.Rows[5].Cells[4], data.EquipmentForm ?? "—");

                // 新增内容
                SetTableCellText(resultTable.Rows[8].Cells[1], data.DownwardElectricalAverageSpeed);
                SetTableCellText(resultTable.Rows[8].Cells[2], data.DownwardElectricalEvaluation);

                SetTableCellText(resultTable.Rows[9].Cells[1], data.DownwardMechanicalAverageSpeed);
                SetTableCellText(resultTable.Rows[9].Cells[2], data.DownwardMechanicalEvaluation);

                SetTableCellText(resultTable.Rows[10].Cells[1], data.UpwardElectricalAverageSpeed);
                SetTableCellText(resultTable.Rows[10].Cells[2], data.UpwardElectricalEvaluation);

                SetTableCellText(resultTable.Rows[11].Cells[1], data.UpwardMechanicalAverageSpeed);
                SetTableCellText(resultTable.Rows[11].Cells[2], data.UpwardMechanicalEvaluation);

                SetTableCellText(resultTable.Rows[12].Cells[1], data.ElectricalSafetyDevice1 ?? "—");
                SetTableCellText(resultTable.Rows[12].Cells[2], data.ElectricalSafetyDevice2 ?? "—");

                // 新增内容
                SetTableCellText(resultTable.Rows[15].Cells[1], data.Notes ?? "-");
                resultTable.Rows[16].Cells[3].Paragraphs[0].Append(data.NextInspectionDate);
                resultTable.Rows[17].Cells[0].Paragraphs[0].Append(data.InspectionDate);
                DateTime dateTime = DateTime.ParseExact(data.InspectionDate, "yyyy年MM月dd日", CultureInfo.InvariantCulture);
                string shenheDate = dateTime.AddDays(1).ToString("yyyy年MM月dd日");
                string pizhunDate = dateTime.AddDays(2).ToString("yyyy年MM月dd日");
                resultTable.Rows[18].Cells[0].Paragraphs[0].Append(shenheDate);
                resultTable.Rows[19].Cells[0].Paragraphs[0].Append(pizhunDate);
            }

            // 保存报告
            var reportDir = Path.Combine(Path.GetDirectoryName(sourcePath), "生成报告", DateTime.Now.ToString("yyyyMM"));
            Directory.CreateDirectory(reportDir);
            var reportName = $"{data.ElevatorRegistrationCode}_生成报告_V{DateTime.Now:HHmmss}.docx";
            doc.SaveAs(Path.Combine(reportDir, reportName));

            ShowColoredMessage($"报告已生成：{Path.Combine(reportDir, reportName)}", ConsoleColor.Green);
        }

        #region 工具方法

        static string ExtractNumber(DocX doc)
        {
            string pattern = @"编号：RTE-JX(\s*)(.*)";
            var paragraphsWithNumber = doc.Paragraphs
                .Where(p => Regex.IsMatch(p.Text, pattern))
                .Select(p => p.Text)
                .FirstOrDefault();

            if (paragraphsWithNumber != null)
            {
                var match = Regex.Match(paragraphsWithNumber, pattern);
                return match.Groups[2].Value.Trim();
            }

            return string.Empty;
        }

        static string GetCellText(Table table, int row, int col)
        {
            return table.Rows[row - 1].Cells[col - 1].Paragraphs.First().Text.Trim();
        }

        static string ConverSpeedText(string input)
        {
            return input == "-" ? "-" : input + "m/s";
        }

        static string ConvertEvaluation(string input)
        {
            return input == "√" ? "合格" : input;
        }

        static string ConvertSafetyDevice(string input)
        {
            return input == "√" ? "符合" : input;
        }

        static void SetTableCellText(Cell cell, string text)
        {
            cell.Paragraphs[0].RemoveText(0, cell.Paragraphs[0].Text.Length);
            cell.Paragraphs[0].Append(text);
        }

        static void ShowColoredMessage(string message, ConsoleColor color)
        {
            Console.ForegroundColor = color;
            Console.WriteLine(message);
            Console.ResetColor();
        }
        #endregion
    }
}