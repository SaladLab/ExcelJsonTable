using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelJsonTable
{
    public partial class RibbonMenu
    {
        private void RibbonMenu_Load(object sender, RibbonUIEventArgs e)
        {
        }

        private string GetPath(string documentPath, string specifiedPath)
        {
            if (Path.IsPathRooted(specifiedPath))
                return specifiedPath;
            else
                return Path.Combine(documentPath, specifiedPath);
        }

        private void buttonCreate_Click(object sender, RibbonControlEventArgs e)
        {
            var activeWorksheet = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet);

            var firstCell = activeWorksheet.Cells[1, 1].Value;
            if (firstCell is string && string.IsNullOrEmpty(firstCell))
            {
                MessageBox.Show("Create action is only for empty sheet", "Create",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            var filePath = "";
            {
                var dlg = new OpenFileDialog();
                dlg.Filter = "Json files (*.json)|*.json|Text files (*.txt)|*.txt";
                dlg.Multiselect = false;
                dlg.Title = "Select a file to import";
                if (dlg.ShowDialog() != DialogResult.OK)
                    return;
                filePath = dlg.FileName;
                activeWorksheet.Cells[1, 1].Value = filePath;
            }

            try
            {
                ExcelJsonLib.DataTransform.Import(
                    ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet),
                    filePath);
            }
            catch (Exception exception)
            {
                MessageBox.Show("Error:\n" + exception, "Create",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            MessageBox.Show("OK!\n" + filePath, "Create",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void buttonImport_Click(object sender, RibbonControlEventArgs e)
        {
            var activeWorksheet = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet);

            var filePath = "";
            var firstCell = activeWorksheet.Cells[1, 1].Value;
            if (firstCell is string)
            {
                filePath = GetPath(Globals.ThisAddIn.Application.ActiveWorkbook.Path, firstCell);

                var result = MessageBox.Show(string.Format("Do you want to import from \n{0}", filePath), "Import",
                    MessageBoxButtons.YesNo);
                if (result != DialogResult.Yes)
                    return;
            }
            else
            {
                var dlg = new OpenFileDialog();
                dlg.Filter = "Json files (*.json)|*.json|Text files (*.txt)|*.txt";
                dlg.Multiselect = false;
                dlg.Title = "Select a file to import";
                if (dlg.ShowDialog() != DialogResult.OK)
                    return;
                filePath = dlg.FileName;
                activeWorksheet.Cells[1, 1].Value = filePath;
            }

            try
            {
                ExcelJsonLib.DataTransform.Import(
                    (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet,
                    filePath);
            }
            catch (Exception exception)
            {
                MessageBox.Show("Error:\n" + exception, "Import",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            MessageBox.Show("OK!\n" + filePath, "Import",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void buttonExport_Click(object sender, RibbonControlEventArgs e)
        {
            var activeWorksheet = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet);

            var filePath = "";
            var firstCell = activeWorksheet.Cells[1, 1].Value;
            if (firstCell is string)
            {
                filePath = GetPath(Globals.ThisAddIn.Application.ActiveWorkbook.Path, firstCell);
            }
            else
            {
                OpenFileDialog dlg = new OpenFileDialog();
                dlg.Filter = "Json files (*.json)|*.json|Text files (*.txt)|*.txt";
                dlg.Multiselect = false;
                dlg.Title = "Select a file to export";
                if (dlg.ShowDialog() != DialogResult.OK)
                    return;
                filePath = dlg.FileName;
            }

            try
            {
                ExcelJsonLib.DataTransform.Export(
                    (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet,
                    filePath);
            }
            catch (Exception exception)
            {
                MessageBox.Show("Error:\n" + exception, "Export",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            MessageBox.Show("OK!\n" + filePath, "Export",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private static T GetAssemblyInfo<T>() where T : class
        {
            return (T) Assembly.GetExecutingAssembly().GetCustomAttributes(typeof (T), false).FirstOrDefault();
        }

        private void buttonAbout_Click(object sender, RibbonControlEventArgs e)
        {
            MessageBox.Show(
                string.Format("{0}\n{1}\n{2}",
                    GetAssemblyInfo<AssemblyTitleAttribute>().Title,
                    GetAssemblyInfo<AssemblyFileVersionAttribute>().Version,
                    GetAssemblyInfo<AssemblyCopyrightAttribute>().Copyright),
                "ExcelJsonTable",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
