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

		private int maxRowsFirstTableInExcel = 11;
		private int maxRowsSecondTableInExcel = 15;
		private int totalRowsTwoTablesInExcel = 26;

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

			// настройка таблицы - шрифт, невозможность изменения ширины столбцов
			dataGridView_DocData.ColumnHeadersDefaultCellStyle.Font = new Font(this.Font.FontFamily, 10, FontStyle.Regular);
			dataGridView_DocData.DefaultCellStyle.Font = new Font(this.Font.FontFamily, 10, FontStyle.Regular);
			dataGridView_DocData.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;

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
				TextBox_TotalSumKopek
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



		/*
		private void ExportToExcel() {
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

			// иначе продолжаем
			// и пытаемся скопировать файл экселя и заполнить его данными
			try {
				// формируем имя файла
				string nameFile = "OP-6_" + TextBox_DocNum.Text + "_" + DateTimePicker_DocDate.Value.Date.Day.ToString() + "_" + DateTimePicker_DocDate.Value.Date.Month.ToString() + "_" + DateTimePicker_DocDate.Value.Date.Year.ToString() + "_" + DateTimePicker_DocDate.Value.Hour.ToString() + "_" + DateTimePicker_DocDate.Value.Minute.ToString() + "_" + DateTimePicker_DocDate.Value.Second.ToString() + ".XLS";

				// создаем копию эксель дока который будем заполнять
				System.IO.File.Copy("OP-6.XLS", nameFile);

				Excel.Application excel = new Excel.Application();
				Excel.Workbook wb = excel.Workbooks.Open(Directory.GetCurrentDirectory() + "/" + nameFile);
				Excel.Worksheet wsh = (Excel.Worksheet)excel.ActiveSheet;

				// номер дока + дата
				wsh.Cells[14, "AO"] = TextBox_DocNum.Text;
				wsh.Cells[14, "AW"] = DateTimePicker_DocDate.Value.Date.Day.ToString() + "." + DateTimePicker_DocDate.Value.Date.Month.ToString() + "." + DateTimePicker_DocDate.Value.Date.Year.ToString();

				// организация + подразделения
				wsh.Cells[6, "A"] = ComboBox_Organization.Text;
				wsh.Cells[8, "A"] = TextBox_Sender.Text;
				//wsh.Cells[10, "A"] = TextBox_Recipient.Text;

				// коды
				wsh.Cells[6, "BQ"] = TextBox_FormOKPO.Text;
				wsh.Cells[11, "BQ"] = TextBox_ActivityOKDP.Text;
				wsh.Cells[12, "BQ"] = TextBox_OperationType.Text;

				// низ эксельки - сумма и количество
				wsh.Cells[66, "J"] = TextBox_TotalQuantityInWords.Text;
				wsh.Cells[68, "J"] = TextBox_TotalSumRubInWords.Text;
				wsh.Cells[68, "BQ"] = TextBox_TotalSumKopek.Text;

				// верх эксельки - лица
				wsh.Cells[15, "T"] = posts[0];
				wsh.Cells[17, "I"] = posts[1];

				wsh.Cells[15, "AC"] = fullNames[0];
				wsh.Cells[17, "AA"] = fullNames[1];
				wsh.Cells[17, "BK"] = fullNames[2];

				// низ эксельки - лица
				wsh.Cells[70, "J"] = posts[2];
				wsh.Cells[70, "AV"] = posts[3];
				wsh.Cells[72, "J"] = posts[4];

				wsh.Cells[70, "AD"] = fullNames[3];
				wsh.Cells[70, "BL"] = fullNames[4];
				wsh.Cells[72, "AD"] = fullNames[5];

				// таблица "всего по документу"
				wsh.Cells[62, "Y"] = TextBox_TotalAmount.Text;
				wsh.Cells[62, "AB"] = TextBox_TotalAmountLeaveTime_2.Text;
				wsh.Cells[62, "AE"] = TextBox_TotalFactSum.Text;
				wsh.Cells[62, "AH"] = TextBox_TotalAmountLeaveTime_4.Text;
				wsh.Cells[62, "AK"] = TextBox_TotalRecordSum.Text;
				wsh.Cells[62, "AN"] = TextBox_TotalAmountLeaveTime_6.Text;
				wsh.Cells[62, "AQ"] = TextBox_TotalProductsReturned.Text;
				wsh.Cells[62, "AV"] = TextBox_TotalProductsLeavedWithoutReturned.Text;
				wsh.Cells[62, "BE"] = TextBox_TotalSum_DiscountPrice.Text;
				wsh.Cells[62, "BO"] = TextBox_TotalSum_SellingPrice.Text;

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

				FillTableInExcel(wsh, rows, numRowsFirstTable, 26, 37, 0);

				if (numRowsSecondTable > 0) FillTableInExcel(wsh, rows, numRowsSecondTable, 46, 61, maxRowsFirstTableInExcel);


				wb.Save();
				excel.Visible = true;
			}
			catch (Exception e) {
				MessageBox.Show(e.Message);
			}
		}

		private void ClearForm() {
			// очистить все текстбоксы + таблицу
			ClearTable();

			// текстбоксы с плейсхолдерами - пустая строка и установить плейсхолдер
			foreach(var item in placeHolderTextBoxes) {
				item.Text = "";
				item.setPlaceholder();
			}

			// обычные текстбоксы - пустая строка
			foreach(var item in usualTextBoxes) item.Text = "";

			// выпадашка с организациями - индекс -1 чтоб пустой была
			ComboBox_Organization.SelectedIndex = -1;

			// дата - текущая
			DateTimePicker_DocDate.Value = DateTime.Now;

			// форма с ответственными лицами
			personsForm.ClearForm();
		}

		private void ClearTable() {
			// очистить таблицу
			dataGridView_DocData.Rows.Clear();

			// обычные текстбоксы забить пустой строкой
			foreach (var item in usualTextBoxes) item.Text = "";

			// с плейсхолдерами - пустой строкой и восстановить плейсхолдера
			TextBox_TotalQuantityInWords.Text = "";
			TextBox_TotalQuantityInWords.setPlaceholder();
			TextBox_TotalSumRubInWords.Text = "";
			TextBox_TotalSumRubInWords.setPlaceholder();
		}

		private void ToolStripMenuItem_ExportToExcel_Click(object sender, EventArgs e) {
			ExportToExcel();
		}

		private void Button_ExportToExcel_Click(object sender, EventArgs e) {
			ExportToExcel();
		}

		private void ToolStripMenuItem_ClearForm_Click(object sender, EventArgs e) {

			DialogResult dialogResult = MessageBox.Show("Очистить всю форму, включая таблицу?", "Предупреждение", MessageBoxButtons.YesNo);
			if (dialogResult == DialogResult.Yes) {
				ClearForm();
			}
			else if (dialogResult == DialogResult.No) {
				return;
			}
		}

		private void ToolStripMenuItem_ClearTable_Click(object sender, EventArgs e) {

			DialogResult dialogResult = MessageBox.Show("Очистить таблицу на форме?", "Предупреждение", MessageBoxButtons.YesNo);
			if (dialogResult == DialogResult.Yes) {
				ClearTable();
			}
			else if (dialogResult == DialogResult.No) {
				return;
			}
			
		}

		*/


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

			}

		}

		private void dataGridView_DocData_CellValueChanged(object sender, DataGridViewCellEventArgs e)
		{
			// при изменении значения в таблице в определенных столбцах делать пересчет значений

			if (e.RowIndex == -1) return;

			if (e.ColumnIndex >= 1 && e.ColumnIndex <= 4)
			{
				ReCountRow(e.RowIndex);
			}
		}


        
		private void ReCountRow(int row) {

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
				for (int i = 0; i < numRows; i++)
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
				TotalValuesToTextBoxes(TotalProductsLeavedWithoutReturned, TotalSum_SellingPrice);
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


		private void TotalValuesToTextBoxes(double quantity, double sum) {
			// отключить плейсхолдера
			// вывести значение переведенное в строку
			TextBox_TotalSumRubInWords.removePlaceHolder();
			TextBox_TotalQuantityInWords.Text = NumToWord.Translate(quantity);

			// отключить плейсхолдера
			// вывести значение переведенное в строку
			TextBox_TotalSumRubInWords.removePlaceHolder();
			TextBox_TotalSumRubInWords.Text = NumToWord.Translate(Math.Floor(sum));

			// получить значение после запятой и вывести его
			TextBox_TotalSumKopek.Text = (Math.Round( 100 * (sum - Math.Floor(sum)))).ToString();
		}

			private bool CheckData(List<string> posts, List<string> fullNames, List<RowInTable> rows) {

				// посчитать кол-во пустых полей
				// их наименования
				// через messagebox узнать у пользователя

				int numWarnings = 0;
				string warnings = "Не были заполнены следующие поля:\r\n";

				void Checking(string text, string warning) {
					if (string.IsNullOrEmpty(text)) { warnings += (warning + "\r\n"); numWarnings++; }
				}

				Checking(TextBox_DocNum.Text, "Номер документ");

				Checking(TextBox_FormOKPO.Text, "Форма ОКПО");
				Checking(TextBox_ActivityOKDP.Text, "Вид деятельности по ОКДП");
				Checking(TextBox_OperationType.Text, "Вид операции");

				Checking(TextBox_Sender.Text, "Структурное подразделение «отправитель»");
				//Checking(TextBox_Recipient.Text, "Структурное подразделение «получатель»");

				Checking(TextBox_TotalQuantityInWords.Text, "Количество прописью");
				Checking(TextBox_TotalSumRubInWords.Text, "Сумма прописью");

				Checking(posts[0], "Мат. ответств. лицо - должность");
				Checking(posts[1], "Руководитель - должность");
				Checking(posts[2], "Отпустил - должность");
				Checking(posts[3], "Принял - должность");
				Checking(posts[4], "Проверил - должность");

				Checking(fullNames[0], "Мат. ответств. лицо - ФИО");
				Checking(fullNames[1], "Руководитель - ФИО");
				Checking(fullNames[2], "Главный (старший) бухгалтер - ФИО");
				Checking(fullNames[3], "Отпустил - ФИО");
				Checking(fullNames[4], "Принял - ФИО");
				Checking(fullNames[5], "Проверил - ФИО");

				if (rows.Count > totalRowsTwoTablesInExcel) { warnings += "Количество строк в таблице больше 26. Будет записано только 26 строк.\r\n"; numWarnings++; }

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

			private List<RowInTable> GetAllRows() {
				// получить список всех строк таблицы формы

				List<RowInTable> res = new List<RowInTable>();

				int numRows = dataGridView_DocData.Rows.Count;
				for (int i = 0; i < numRows; i++) {
					RowInTable curRow = new RowInTable();
					curRow.productName = Convert.ToString(dataGridView_DocData[0, i].Value);
					curRow.productCode = Convert.ToString(dataGridView_DocData[1, i].Value);

					curRow.measurmentUnitName = Convert.ToString(dataGridView_DocData[2, i].Value);
					curRow.measurmentUnitCode = Convert.ToString(dataGridView_DocData[3, i].Value);

					curRow.leaveTime_1 = Convert.ToString(dataGridView_DocData[4, i].Value);
					curRow.leaveTime_2 = Convert.ToString(dataGridView_DocData[5, i].Value); 
					curRow.leaveTime_3 = Convert.ToString(dataGridView_DocData[6, i].Value);
					curRow.leaveTime_4 = Convert.ToString(dataGridView_DocData[7, i].Value);
					curRow.leaveTime_5 = Convert.ToString(dataGridView_DocData[8, i].Value);
					curRow.leaveTime_6 = Convert.ToString(dataGridView_DocData[9, i].Value);

					curRow.productReturned = Convert.ToString(dataGridView_DocData[10, i].Value);
					curRow.productReturnedWithoutLeaving = Convert.ToString(dataGridView_DocData[11, i].Value);

					curRow.discountPrice = Convert.ToString(dataGridView_DocData[12, i].Value);
					curRow.discountPriceSum = Convert.ToString(dataGridView_DocData[13, i].Value);

					curRow.sellingPrice = Convert.ToString(dataGridView_DocData[14, i].Value);
					curRow.sellingPriceSum = Convert.ToString(dataGridView_DocData[15, i].Value);

					curRow.note = Convert.ToString(dataGridView_DocData[16, i].Value);

					res.Add(curRow);
				}

				return res;
			}


			private void FillTableInExcel(Excel.Worksheet wsh, List<RowInTable> rows, int numRows, int startRowIndex, int totalRowIndex, int listShift) {
				// заполнить таблицу в экселе

				double[] temp = { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };

				// заполнение таблицы и вычисление значений для строки "итого"
				for (int i = 0; i < numRows; i++) {
					RowInTable curRow = rows[listShift + i];

					wsh.Cells[startRowIndex + i, "A"] = curRow.productName;
					wsh.Cells[startRowIndex + i, "M"] = curRow.productCode;

					wsh.Cells[startRowIndex + i, "Q"] = curRow.measurmentUnitName;
					wsh.Cells[startRowIndex + i, "U"] = curRow.measurmentUnitCode;

					wsh.Cells[startRowIndex + i, "Y"] = curRow.leaveTime_1;
					wsh.Cells[startRowIndex + i, "AB"] = curRow.leaveTime_2;
					wsh.Cells[startRowIndex + i, "AE"] = curRow.leaveTime_3;
					wsh.Cells[startRowIndex + i, "AH"] = curRow.leaveTime_4;
					wsh.Cells[startRowIndex + i, "AK"] = curRow.leaveTime_5;
					wsh.Cells[startRowIndex + i, "AN"] = curRow.leaveTime_6;

					wsh.Cells[startRowIndex + i, "AQ"] = curRow.productReturned;
					wsh.Cells[startRowIndex + i, "AV"] = curRow.productReturnedWithoutLeaving;

					wsh.Cells[startRowIndex + i, "AZ"] = curRow.discountPrice;
					wsh.Cells[startRowIndex + i, "BE"] = curRow.discountPriceSum;

					wsh.Cells[startRowIndex + i, "BJ"] = curRow.sellingPrice;
					wsh.Cells[startRowIndex + i, "BO"] = curRow.sellingPriceSum;

					wsh.Cells[startRowIndex + i, "BT"] = curRow.note;

					temp[0] += (curRow.leaveTime_1 == string.Empty ? 0.0 : Convert.ToDouble(curRow.leaveTime_1));
					temp[1] += (curRow.leaveTime_2 == string.Empty ? 0.0 : Convert.ToDouble(curRow.leaveTime_2));
					temp[2] += (curRow.leaveTime_3 == string.Empty ? 0.0 : Convert.ToDouble(curRow.leaveTime_3));
					temp[3] += (curRow.leaveTime_4 == string.Empty ? 0.0 : Convert.ToDouble(curRow.leaveTime_4));
					temp[4] += (curRow.leaveTime_5 == string.Empty ? 0.0 : Convert.ToDouble(curRow.leaveTime_5));
					temp[5] += (curRow.leaveTime_6 == string.Empty ? 0.0 : Convert.ToDouble(curRow.leaveTime_6));
					temp[6] += (curRow.productReturned == string.Empty ? 0.0 : Convert.ToDouble(curRow.productReturned));
					temp[7] += (curRow.productReturnedWithoutLeaving == string.Empty ? 0.0 : Convert.ToDouble(curRow.productReturnedWithoutLeaving));
					temp[8] += (curRow.discountPriceSum == string.Empty ? 0.0 : Convert.ToDouble(curRow.discountPriceSum));
					temp[9] += (curRow.sellingPriceSum == string.Empty ? 0.0 : Convert.ToDouble(curRow.sellingPriceSum));
				}

				// "итого" для таблицы
				wsh.Cells[totalRowIndex, "Y"] = temp[0];
				wsh.Cells[totalRowIndex, "AB"] = temp[1];
				wsh.Cells[totalRowIndex, "AE"] = temp[2];
				wsh.Cells[totalRowIndex, "AH"] = temp[3];
				wsh.Cells[totalRowIndex, "AK"] = temp[4];
				wsh.Cells[totalRowIndex, "AN"] = temp[5];
				wsh.Cells[totalRowIndex, "AQ"] = temp[6];
				wsh.Cells[totalRowIndex, "AV"] = temp[7];
				wsh.Cells[totalRowIndex, "BE"] = temp[8];
				wsh.Cells[totalRowIndex, "BO"] = temp[9];
			}

			private void dataGridView_DocData_UserAddedRow(object sender, DataGridViewRowEventArgs e) {
				if (dataGridView_DocData.Rows.Count > 26) dataGridView_DocData.AllowUserToAddRows = false; else dataGridView_DocData.AllowUserToDeleteRows = true;

				ReCountTotal();
			}

			private void dataGridView_DocData_UserDeletedRow(object sender, DataGridViewRowEventArgs e) {
				if (dataGridView_DocData.Rows.Count < 26) dataGridView_DocData.AllowUserToAddRows = true;

				ReCountTotal();
			}

			private void ToolStripMenuItem_Exit_Click(object sender, EventArgs e) {
				Close();
			}

			private void panel_table_Paint(object sender, PaintEventArgs e)
			{

			}
		}*/
    }
}
	