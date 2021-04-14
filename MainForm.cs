using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;

namespace Catering_OP_6 {
	public partial class MainForm : Form
	{

		private int maxRowsFirstTableInExcel = 29;
		private int maxRowsSecondTableInExcel = 24;
		private int totalRowsTwoTablesInExcel = 53;

		PersonsForm personsForm;

		List<PlaceHolderTextBox> placeHolderTextBoxes;
		List<TextBox> usualTextBoxes;

		private List<string> organiztions = new List<string>() {
			"ООО \"Едим как дома\"",
			"ООО \"Пельмешки\"",
			"ООО \"Кушать подано\""
		};

		private List<string> structPodrazd = new List<string>() {
			"Столовая №1",
			"Столовая №2",
			"Столовая №3",
		};

		private List<string> productsName = new List<string>() {
			"Борщ \"Киевский\"",
			"Солянка",
			"Салат \"Цезарь\"",
			"Гречка с грибами"
		};

		private List<string> productsCode = new List<string>() {
			"6534",
			"4534",
			"4623",
			"5678"
		};

		private List<string> cardNumber = new List<string>() {
			"123",
			"126",
			"127",
			"129"
		};


		private List<string> factPrice = new List<string>() {
			"50,90",
			"57,50",
			"60",
			"74"
		};

		private List<string> recordPrice = new List<string>() {
			"45",
			"52",
			"54",
			"68"
		};


		public MainForm()
		{
			// инициализация формы со всеми объектами
			InitializeComponent();

			this.textBox_salt.Controls[0].Visible = false;
			this.textBox_spices.Controls[0].Visible = false;

			// настройка таблицы - шрифт, невозможность изменения ширины столбцов
			dataGridView_DocData.ColumnHeadersDefaultCellStyle.Font = new Font(this.Font.FontFamily, 10, FontStyle.Regular);
			dataGridView_DocData.DefaultCellStyle.Font = new Font(this.Font.FontFamily, 10, FontStyle.Regular);
			//dataGridView_DocData.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;

			// создание формы "ответственные лица"
			personsForm = new PersonsForm(this);

			// текстбоксы с плейсхолдерами данной формы
			placeHolderTextBoxes = new List<PlaceHolderTextBox>() {
				TextBox_DocNum,
				TextBox_ActivityOKDP,
				TextBox_FormOKPO,
				TextBox_OperationType,
				TextBox_TotalSumRubInWords
			};

			// обычные текстбоксы данной формы
			usualTextBoxes = new List<TextBox>() {
				TextBox_TotalAmount,
				TextBox_TotalFactSum,
				TextBox_TotalRecordSum,
				TextBox_TotalSumKopek,
				textBox_salt_cop,
				textBox_salt_rub,
				textBox_spices_cop,
				textBox_spices_rub,
				textBox_total_spices_salt_rub,
				textBox_total_spices_salt_cop
			};

			// в таблицу добавить в столбцы значения для выборов
			DataGridViewColumnCollection columns = dataGridView_DocData.Columns;
			DataGridViewComboBoxColumn columnCardNumber = (DataGridViewComboBoxColumn)columns[1];
			DataGridViewComboBoxColumn columnProductsName = (DataGridViewComboBoxColumn)columns[2];
			DataGridViewComboBoxColumn columnCode = (DataGridViewComboBoxColumn)columns[3];

			// заполняем выпадашки в таблице
			foreach (var item in cardNumber) columnCardNumber.Items.Add(item);
			foreach (var item in productsName) columnProductsName.Items.Add(item);
			foreach (var item in productsCode) columnCode.Items.Add(item);




			// выпадашка "организации" и структурное подразделение
			foreach (var item in organiztions) ComboBox_Organization.Items.Add(item);
			foreach (var item in structPodrazd) comboBox_StructPodrazd.Items.Add(item);
		}

		private void Link_ResponsiblePersons_Click(object sender, LinkLabelLinkClickedEventArgs e)
		{
			if (personsForm.Visible) personsForm.Hide(); else personsForm.Show();
		}


		private void ExportToExcel() 
		{
			// создание эксель файла и загрузка по ячейкам данных

			// возможно сделать проверку данных

			// с формы "ответственные лица" собираем инфу
			List<string> posts = personsForm.GetPosts();
			List<string> fullNames = personsForm.GetFullNames();

			// из таблицы на форме собираем все строки
			List<RowInTable> rows = GetAllRows();

			// делаем проверку, что поля заполнены
			// и таблица не пустая или что в ней больше 20 записей
			// если пользователь отказывается - то выход из экспорта
			if (CheckData(posts, fullNames, rows) == false) return;

			for (int i = 0; i < rows.Count - 1; i++)
				if (rows[i].checkRow() == false)
                {
					MessageBox.Show("Строка " + rows[i].row_num.ToString() + " заполнена некорректно! (Имеются пустые поля)", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Warning);
					return;
                }

			// иначе продолжаем
			// и пытаемся скопировать файл экселя и заполнить его данными
			try {
				// формируем имя файла
				string nameFile = "OP-12_" + TextBox_DocNum.Text + "_" + DateTimePicker_DocDate.Value.Date.Day.ToString() + "_" + DateTimePicker_DocDate.Value.Date.Month.ToString() + "_" + DateTimePicker_DocDate.Value.Date.Year.ToString() + "_" + DateTimePicker_DocDate.Value.Hour.ToString() + "_" + DateTimePicker_DocDate.Value.Minute.ToString() + "_" + DateTimePicker_DocDate.Value.Second.ToString() + ".XLS";

				// создаем копию эксель дока который будем заполнять
				System.IO.File.Copy("OP-12.xls", nameFile);

				Excel.Application excel = new Excel.Application();
				Excel.Workbook wb = excel.Workbooks.Open(Directory.GetCurrentDirectory() + "/" + nameFile);
				Excel.Worksheet wsh = (Excel.Worksheet)excel.ActiveSheet;

				// номер дока + дата
				wsh.Cells[14, "U"] = TextBox_DocNum.Text;
				wsh.Cells[14, "AB"] = DateTimePicker_DocDate.Value.Date.Day.ToString() + "." + DateTimePicker_DocDate.Value.Date.Month.ToString() + "." + DateTimePicker_DocDate.Value.Date.Year.ToString();
				wsh.Cells[17, "AL"] = DateTimePicker_DocDate.Value.Date.Day.ToString();
				wsh.Cells[17, "AN"] = DateTimePicker_DocDate.Value.Date.Month.ToString();
				wsh.Cells[17, "AU"] = DateTimePicker_DocDate.Value.Date.Year.ToString();

				// организация + подразделения
				wsh.Cells[6, "A"] = ComboBox_Organization.Text;
				wsh.Cells[8, "A"] = comboBox_StructPodrazd.Text;

				// коды
				wsh.Cells[6, "AO"] = TextBox_FormOKPO.Text;
				wsh.Cells[9, "AO"] = TextBox_ActivityOKDP.Text;
				wsh.Cells[10, "AO"] = TextBox_OperationType.Text;

				
				// лица
				wsh.Cells[13, "AM"] = posts[0];

				wsh.Cells[100, "AC"] = fullNames[0];
				wsh.Cells[102, "AC"] = fullNames[1];
				wsh.Cells[104, "AC"] = fullNames[2];
				wsh.Cells[111, "O"] = fullNames[3];
				wsh.Cells[113, "U"] = fullNames[4];

				wsh.Cells[102, "H"] = posts[2];
				wsh.Cells[104, "H"] = posts[3];


				// специи и соль
				wsh.Cells[93, "E"] = textBox_spices.Text;
				wsh.Cells[93, "T"] = textBox_spices_rub.Text;
				wsh.Cells[93, "AL"] = textBox_spices_cop.Text;

				wsh.Cells[95, "D"] = textBox_salt.Text;
				wsh.Cells[95, "T"] = textBox_salt_rub.Text;
				wsh.Cells[95, "AL"] = textBox_salt_cop.Text;

				// специи и соль - итого
				wsh.Cells[97, "T"] = textBox_total_spices_salt_rub.Text;
				wsh.Cells[97, "AL"] = textBox_total_spices_salt_cop.Text;

				// выручка кассы
				wsh.Cells[107, "I"] = TextBox_TotalSumRubInWords.Text; 
				wsh.Cells[109, "AS"] = TextBox_TotalSumKopek.Text; 


				// итог по таблице
				wsh.Cells[56, "V"] = TextBox_TotalAmount.Text;
				wsh.Cells[56, "AE"] = TextBox_TotalFactSum.Text;
				wsh.Cells[56, "AO"] = TextBox_TotalRecordSum.Text;

				wsh.Cells[90, "V"] = 0.ToString();
				wsh.Cells[90, "AE"] = 0.ToString();
				wsh.Cells[90, "AO"] = 0.ToString();

				wsh.Cells[91, "V"] = TextBox_TotalAmount.Text;
				wsh.Cells[91, "AE"] = TextBox_TotalFactSum.Text;
				wsh.Cells[91, "AO"] = TextBox_TotalRecordSum.Text;


				// таблицы 1 и 2
				int numRowsFirstTable, numRowsSecondTable;

				if (rows.Count > maxRowsFirstTableInExcel) {
					numRowsFirstTable = maxRowsFirstTableInExcel;

					if (rows.Count < totalRowsTwoTablesInExcel)
						numRowsSecondTable = rows.Count - numRowsFirstTable;
					else
						numRowsSecondTable = maxRowsSecondTableInExcel;
				}
				else {
					numRowsFirstTable = rows.Count;
					numRowsSecondTable = 0;
				}

				FillTableInExcel(wsh, rows, numRowsFirstTable, 26, 56, 0);

				if (numRowsSecondTable > 0) FillTableInExcel(wsh, rows, numRowsSecondTable, 65, 90, maxRowsFirstTableInExcel);


				wb.Save();
				excel.Visible = true;
			}
			catch (Exception e) {
				MessageBox.Show(e.Message);
			}
		}

		private void ClearForm()
		{
			// очистить все текстбоксы + таблицу
			ClearTable();

			// текстбоксы с плейсхолдерами - пустая строка и установить плейсхолдер
			foreach (var item in placeHolderTextBoxes)
			{
				item.Text = "";
				item.setPlaceholder();
			}



			// обычные текстбоксы - пустая строка
			foreach (var item in usualTextBoxes) item.Text = "";

			// выпадашка с организациями - индекс -1 чтоб пустой была
			ComboBox_Organization.SelectedIndex = -1;

			// дата - текущая
			DateTimePicker_DocDate.Value = DateTime.Now;

			// форма с ответственными лицами
			personsForm.ClearForm();
		}

		private void ClearTable()
		{
			// очистить таблицу
			dataGridView_DocData.Rows.Clear();

			// обычные текстбоксы забить пустой строкой
			foreach (var item in usualTextBoxes) item.Text = "";

			textBox_salt.Value = 0;
			textBox_spices.Value = 0;

			// с плейсхолдерами - пустой строкой и восстановить плейсхолдера
			TextBox_TotalSumRubInWords.Text = "";
			TextBox_TotalSumRubInWords.setPlaceholder();
		}

		private void ToolStripMenuItem_ExportToExcel_Click(object sender, EventArgs e)
		{
			ExportToExcel();
		}


		private void ToolStripMenuItem_ClearForm_Click(object sender, EventArgs e)
		{

			DialogResult dialogResult = MessageBox.Show("Очистить всю форму, включая таблицу?", "Предупреждение", MessageBoxButtons.YesNo);
			if (dialogResult == DialogResult.Yes)
			{
				ClearForm();
			}
			else if (dialogResult == DialogResult.No)
			{
				return;
			}
		}

		private void ToolStripMenuItem_ClearTable_Click(object sender, EventArgs e)
		{

			DialogResult dialogResult = MessageBox.Show("Очистить таблицу на форме?", "Предупреждение", MessageBoxButtons.YesNo);
			if (dialogResult == DialogResult.Yes)
			{
				ClearTable();
			}
			else if (dialogResult == DialogResult.No)
			{
				return;
			}

		}



		private void dataGridView_DocData_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e) 
		{

			if (dataGridView_DocData.CurrentCell.ColumnIndex >= 1 && dataGridView_DocData.CurrentCell.ColumnIndex <= 3)
			{
				ComboBox combo = e.Control as ComboBox;
				combo.SelectedIndexChanged -= new EventHandler(Control_Changed);
				combo.SelectedIndexChanged += new EventHandler(Control_Changed);
			}
			else if (dataGridView_DocData.CurrentCell.ColumnIndex == 4)
            {
				TextBox tb = (TextBox)e.Control;
				tb.KeyPress += new KeyPressEventHandler(tb_KeyPress);

			}

		}

		void tb_KeyPress(object sender, KeyPressEventArgs e)
		{
			if (!(Char.IsDigit(e.KeyChar)))
			{
				if (e.KeyChar != (char)Keys.Back)
				{ e.Handled = true; }
			}
		}


		private void Control_Changed(object sender, System.EventArgs e)
		{
			// загрузка в столбцы информации связанной с выбором в выпадашках таблицы

			int col = dataGridView_DocData.CurrentCell.ColumnIndex;
			

			if (col >= 1 && col <= 3)
            {
				int i = ((ComboBox)sender).SelectedIndex;
				if (col == 1)
                {
					// выбрали карту
					// заменяем имя
					dataGridView_DocData.CurrentRow.Cells[2].Value = productsName[i];

					// заменяем код
					dataGridView_DocData.CurrentRow.Cells[3].Value = productsCode[i];
				}
				else if (col == 2)
                {
					//выбрали имя
					// заменяем карту
					dataGridView_DocData.CurrentRow.Cells[1].Value = cardNumber[i];

					// заменяем код
					dataGridView_DocData.CurrentRow.Cells[3].Value = productsCode[i];
				}
				else
				{
					//выбрали код
					// заменяем карту
					dataGridView_DocData.CurrentRow.Cells[1].Value = cardNumber[i];

					// заменяем имя
					dataGridView_DocData.CurrentRow.Cells[2].Value = productsName[i];
				}

				// факт цена за единицу
				dataGridView_DocData.CurrentRow.Cells[5].Value = factPrice[i];

				// произв. цена за единицу
				dataGridView_DocData.CurrentRow.Cells[7].Value = recordPrice[i];

				ReCountRow(dataGridView_DocData.CurrentRow.Index);
				updateValues();

			}

		}

		private void dataGridView_DocData_CellValueChanged(object sender, DataGridViewCellEventArgs e)
		{
			// при изменении значения в таблице в определенных столбцах делать пересчет значений

			if (e.RowIndex == -1) return;

			if (e.ColumnIndex >= 1 && e.ColumnIndex <= 4)
			{
				ReCountRow(e.RowIndex);
				updateValues();
			}
		}


        
		private void ReCountRow(int row) 
		{
			if (String.IsNullOrWhiteSpace(Convert.ToString(dataGridView_DocData[4, row].Value))
				|| String.IsNullOrWhiteSpace(Convert.ToString(dataGridView_DocData[5, row].Value))
				|| String.IsNullOrWhiteSpace(Convert.ToString(dataGridView_DocData[7, row].Value)))
				return;

			if (String.IsNullOrWhiteSpace(Convert.ToString(dataGridView_DocData[5, row].Value))
				|| String.IsNullOrWhiteSpace(Convert.ToString(dataGridView_DocData[7, row].Value)))
				return;

			// если будет ошибка - например ввели букву, то проставить минусы
			// чтоб прога не вылетела
			try {

				// суммы исходя из прайса и количества
				dataGridView_DocData[6, row].Value = (Convert.ToDouble(dataGridView_DocData[5, row].Value) * Convert.ToDouble(dataGridView_DocData[4, row].Value)).ToString();
				dataGridView_DocData[8, row].Value = (Convert.ToDouble(dataGridView_DocData[7, row].Value) * Convert.ToDouble(dataGridView_DocData[4, row].Value)).ToString();

				// пересчитать для текстбоксов "итог" под таблицей
				ReCountTotal();
			}
			catch (Exception e) {
				// сообщение об ошибке
				MessageBox.Show(e.Message);

				// в данных столбцах нельзя вывести нормальное значение
				// заменяем минусами
				dataGridView_DocData[6, row].Value = "-";
				dataGridView_DocData[8, row].Value = "-";

			}
		}


		
		private void ReCountTotal() {
			// посчитать значения для строки "итого"

			try {
				int numRows = dataGridView_DocData.Rows.Count;

				double TotalSum_Fact= 0, TotalSum_Record = 0, Total_Amount = 0;

				// суммируем во всех строках таблицы нужные значения
				for (int i = 0; i < numRows - 1; i++)
				{ 
					Total_Amount += Convert.ToInt32(dataGridView_DocData[4, i].Value);
					TotalSum_Fact += Convert.ToDouble(dataGridView_DocData[6, i].Value);
					TotalSum_Record += Convert.ToDouble(dataGridView_DocData[8, i].Value);				
				}

				// выводим в текстбоксы под таблицей в "итог"
				TextBox_TotalAmount.Text = Total_Amount.ToString();
				TextBox_TotalFactSum.Text = TotalSum_Fact.ToString();
				TextBox_TotalRecordSum.Text = TotalSum_Record.ToString();
				
				// функция чтобы вывести в текстбоксы словами значения
				TotalValuesToTextBoxes(Math.Floor(TotalSum_Fact));
				TextBox_TotalSumKopek.Text = (Math.Round(100 * (TotalSum_Fact - Math.Floor(TotalSum_Fact)))).ToString();
			}
			catch (Exception e) 
			{
				MessageBox.Show(e.Message);
			}
		}

        private void dataGridView_DocData_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)
        {
			int index = e.RowIndex;
			string indexStr = (index + 1).ToString();
			this.dataGridView_DocData.Rows[index].Cells[0].Value = indexStr;
		}


		private void TotalValuesToTextBoxes(double sum) {
			// отключить плейсхолдера
			// вывести значение переведенное в строку
			TextBox_TotalSumRubInWords.removePlaceHolder();
			TextBox_TotalSumRubInWords.Text = NumToWord.Translate(sum);

		}

        private void Button_ExportToExcel_Click(object sender, EventArgs e)
        {
			ExportToExcel();
		}

    
		private bool CheckData(List<string> posts, List<string> fullNames, List<RowInTable> rows) 
		{

			// посчитать кол-во пустых полей
			// их наименования
			// через messagebox узнать у пользователя

			int numWarnings = 0;
			string warnings = "Не были заполнены следующие поля:\r\n";

			void Checking(string text, string warning) {
				if (string.IsNullOrEmpty(text)) { warnings += (warning + "\r\n"); numWarnings++; }
			}

			Checking(TextBox_DocNum.Text, "Номер документа");
			Checking(TextBox_FormOKPO.Text, "Форма ОКПО");
			Checking(TextBox_ActivityOKDP.Text, "Вид деятельности по ОКДП");
			Checking(TextBox_OperationType.Text, "Вид операции");
			Checking(comboBox_StructPodrazd.Text, "Структурное подразделение");
			Checking(TextBox_TotalSumRubInWords.Text, "Сумма прописью");
			Checking(this.textBox_spices.Text, "% специй");
			Checking(this.textBox_salt.Text, "% соли");


			Checking(posts[0], "Утвердил - должность");

			Checking(posts[1], "Член комиссии 1 - должность");
			Checking(posts[2], "Член комиссии 2 - должность");
			Checking(posts[3], "Член комиссии 3 - должность");

			Checking(fullNames[0], "Член комиссии 1 - ФИО");
			Checking(fullNames[1], "Член комиссии 2 - ФИО");
			Checking(fullNames[2], "Член комиссии 3 - ФИО");
			Checking(fullNames[3], "Кассир - ФИО");
			Checking(fullNames[4], "Проверил - ФИО");

			if (rows.Count > totalRowsTwoTablesInExcel) { warnings += "Количество строк в таблице больше 53. Будет записано только 53 строки.\r\n"; numWarnings++; }

			if (numWarnings > 0) {
				warnings += "\r\nВсего предупреждений: " + numWarnings + ". Продолжить?";

				DialogResult dialogResult = MessageBox.Show( warnings, "Предупреждение", MessageBoxButtons.YesNo);
				if (dialogResult == DialogResult.Yes) {
					return true;
				}
				else if (dialogResult == DialogResult.No) {
					return false;
				}
			}

			// по дефолту - все ок
			return true;
		}


		private List<RowInTable> GetAllRows() 
		{
			// получить список всех строк таблицы формы

			List<RowInTable> res = new List<RowInTable>();

			int numRows = dataGridView_DocData.Rows.Count;
			for (int i = 0; i < numRows; i++) 
			{
				RowInTable curRow = new RowInTable();
				curRow.row_num = Convert.ToString(dataGridView_DocData[0, i].Value);
				curRow.card_number = Convert.ToString(dataGridView_DocData[1, i].Value);
				curRow.name = Convert.ToString(dataGridView_DocData[2, i].Value);
				curRow.code = Convert.ToString(dataGridView_DocData[3, i].Value);
				curRow.amount = Convert.ToString(dataGridView_DocData[4, i].Value);
				curRow.fact_price = Convert.ToString(dataGridView_DocData[5, i].Value);
				curRow.fact_sum = Convert.ToString(dataGridView_DocData[6, i].Value);
				curRow.record_price = Convert.ToString(dataGridView_DocData[7, i].Value);
				curRow.record_sum = Convert.ToString(dataGridView_DocData[8, i].Value);
				curRow.note = Convert.ToString(dataGridView_DocData[9, i].Value);
				
				res.Add(curRow);
			}

			return res;
		}


		private void FillTableInExcel(Excel.Worksheet wsh, List<RowInTable> rows, int numRows, int startRowIndex, int totalRowIndex, int listShift)
		{
			// заполнить таблицу в экселе

			int total_amount = 0;
			double total_fact = 0;
			double total_record = 0;

			// заполнение таблицы и вычисление значений для строки "итого"
			for (int i = 0; i < numRows - 1; i++) {
				RowInTable curRow = rows[listShift + i];

				wsh.Cells[startRowIndex + i, "A"] = curRow.row_num;

				wsh.Cells[startRowIndex + i, "D"] = curRow.card_number;
				wsh.Cells[startRowIndex + i, "H"] = curRow.name;
				wsh.Cells[startRowIndex + i, "S"] = curRow.code;
				wsh.Cells[startRowIndex + i, "V"] = curRow.amount;

				wsh.Cells[startRowIndex + i, "Z"] = curRow.fact_price;
				wsh.Cells[startRowIndex + i, "AE"] = curRow.fact_sum;
				wsh.Cells[startRowIndex + i, "AJ"] = curRow.record_price;
				wsh.Cells[startRowIndex + i, "AO"] = curRow.record_sum;
				wsh.Cells[startRowIndex + i, "AT"] = curRow.note;

				total_amount += Convert.ToInt32(curRow.amount);
				total_fact += Convert.ToDouble(curRow.fact_sum);
				total_record += Convert.ToDouble(curRow.record_sum);

			}

			// "итого" для таблицы
			wsh.Cells[totalRowIndex, "V"] = total_amount;
			wsh.Cells[totalRowIndex, "AE"] = total_fact;
			wsh.Cells[totalRowIndex, "AO"] = total_record;

		}

		private void dataGridView_DocData_UserAddedRow(object sender, DataGridViewRowEventArgs e) {
			if (dataGridView_DocData.Rows.Count > 53) dataGridView_DocData.AllowUserToAddRows = false; else dataGridView_DocData.AllowUserToDeleteRows = true;

			ReCountTotal();
		}

		private void dataGridView_DocData_UserDeletedRow(object sender, DataGridViewRowEventArgs e) {
			if (dataGridView_DocData.Rows.Count < 53) dataGridView_DocData.AllowUserToAddRows = true;

			ReCountTotal();
		}

		private void ToolStripMenuItem_Exit_Click(object sender, EventArgs e) {
			Close();
		}

		private void panel_table_Paint(object sender, PaintEventArgs e)
		{

		}

        private void textBox_spices_ValueChanged(object sender, EventArgs e)
        {
			updateValues();
        }

        private void textBox_salt_ValueChanged(object sender, EventArgs e)
        {
			updateValues();
		}

		private void updateValues()
        {
			double spices = Convert.ToDouble(textBox_spices.Value) / 100;
			double salt = Convert.ToDouble(textBox_salt.Value) / 100;
			double total = 0;

			double fact_price = 0;
			if (String.IsNullOrWhiteSpace(TextBox_TotalFactSum.Text) == false)
				fact_price = Convert.ToDouble(TextBox_TotalFactSum.Text);

			spices *= fact_price;
			salt *= fact_price;
			total = spices + salt;

			//this.textBox_salt_rub.Text = NumToWord.Translate(Math.Floor(salt));
			//this.textBox_salt_cop.Text = NumToWord.Translate(Math.Round(100 * (salt - Math.Floor(salt))));
			//this.textBox_spices_rub.Text = NumToWord.Translate(Math.Floor(spices));
			//this.textBox_spices_cop.Text = NumToWord.Translate(Math.Round(100 * (spices - Math.Floor(spices))));

			this.textBox_salt_rub.Text = Math.Floor(salt).ToString();
			this.textBox_salt_cop.Text = Math.Round(100 * (salt - Math.Floor(salt))).ToString();
			this.textBox_spices_rub.Text = Math.Floor(spices).ToString();
			this.textBox_spices_cop.Text = Math.Round(100 * (spices - Math.Floor(spices))).ToString();
			this.textBox_total_spices_salt_rub.Text = Math.Floor(total).ToString();
			this.textBox_total_spices_salt_cop.Text = Math.Round(100 * (total - Math.Floor(total))).ToString();

		}
    }
}
	