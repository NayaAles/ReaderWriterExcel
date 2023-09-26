using System.Collections.Generic;

namespace ReaderWriterExcel
{
    public partial class ThisWorkbook
    {
        public ReaderWriterExcel readerWriterExcel = new ReaderWriterExcel();
        public List<InData> inDatas = new List<InData>();

        private void ThisWorkbook_Startup(object sender, System.EventArgs e)
        {
            Main();
        }

        private void ThisWorkbook_Shutdown(object sender, System.EventArgs e)
        {
            ReaderWriterExcel.ExcelKill();
        }

        public void Main()
        {
            readerWriterExcel.ReadFromExcel(inDatas, 0);
            readerWriterExcel.SaveToExcel(inDatas, 1);

            ReaderWriterExcel.ExcelKill();
        }

        #region Код, созданный конструктором VSTO

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisWorkbook_Startup);
            this.Shutdown += new System.EventHandler(ThisWorkbook_Shutdown);
        }

        #endregion

    }
}
